Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports System.Xml
Imports OSGeo.MapGuide

Public Class IFC
    Public Shared Function FilletAt(ByVal pline As Polyline, ByVal index As Integer, ByVal radius As Double) As Integer
        Dim prev As Integer = If(index = 0 AndAlso pline.Closed, pline.NumberOfVertices - 1, index - 1)

        If pline.GetSegmentType(prev) <> SegmentType.Line OrElse pline.GetSegmentType(index) <> SegmentType.Line Then
            Return 0
        End If

        Dim seg1 As LineSegment2d = pline.GetLineSegment2dAt(prev)
        Dim seg2 As LineSegment2d = pline.GetLineSegment2dAt(index)
        Dim vec1 As Vector2d = seg1.StartPoint - seg1.EndPoint
        Dim vec2 As Vector2d = seg2.EndPoint - seg2.StartPoint
        Dim angle As Double = (Math.PI - vec1.GetAngleTo(vec2)) / 2.0
        Dim dist As Double = radius * Math.Tan(angle)

        If dist = 0.0 OrElse dist > seg1.Length OrElse dist > seg2.Length Then
            Return 0
        End If

        Dim pt1 As Point2d = seg1.EndPoint + vec1.GetNormal() * dist
        Dim pt2 As Point2d = seg2.StartPoint + vec2.GetNormal() * dist
        Dim bulge As Double = Math.Tan(angle / 2.0)

        If Clockwise(seg1.StartPoint, seg1.EndPoint, seg2.EndPoint) Then
            bulge = -bulge
        End If

        pline.AddVertexAt(index, pt1, bulge, 0.0, 0.0)
        pline.SetPointAt(index + 1, pt2)
        Return 1
    End Function

    Public Shared Function Clockwise(ByVal p1 As Point2d, ByVal p2 As Point2d, ByVal p3 As Point2d) As Boolean
        Return ((p2.X - p1.X) * (p3.Y - p1.Y) - (p2.Y - p1.Y) * (p3.X - p1.X)) < 0.00000001
    End Function

    '=============================================================================================
    'Создать сетку Face на основе выбранной полилинии
    '=============================================================================================
    <CommandMethod("ImportOSM")>
    Public Sub ImportOSM()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim CodeCoordSys As String = CivilDoc.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
        Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
        Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
        Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
        Dim datumDict As MgCoordinateSystemDatumDictionary = csCatalog.GetDatumDictionary()
        Dim ellipsoidDict As MgCoordinateSystemEllipsoidDictionary = csCatalog.GetEllipsoidDictionary
        Dim LL84Cs As MgCoordinateSystem = csDict.GetCoordinateSystem("LL84")
        Dim DrawingSysCoord As MgCoordinateSystem = csDict.GetCoordinateSystem(CodeCoordSys)

        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            'начальный путь к файлу-шаблону
            Dim PutchTemplateXml As String = ""
            Dim PutchTemplateDwg As String = ""
            Dim nameFileXML As String = "Стандарт_ADSK.xml"
            Dim nameFileDwg As String = "Стандарт_ADSK.dwg"
            Dim boolFindTeplates As Boolean = FuncFiles.FuncFindTemplateFiles(nameFileXML, nameFileDwg, PutchTemplateXml, PutchTemplateDwg)
            '=============================================================================================================
            'проверяем и загружаем все стили из шаблона
            '=============================================================================================================
            '3. Путь к файлам шаблонам с условными знами
            If boolFindTeplates = False Then
                MsgBox("Шаблон не найден. Добавьте файлы Стандарт_ADSK.xml и Стандарт_ADSK.dwg в вспомогательные файлы поддержки. Макро прервано!!!")
                Exit Sub
            End If
            Dim NamePSSetTable As String = "LAND_CADASTRAL_BUILDING"
            Dim boolPSBuild As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, NamePSSetTable)
            Dim IDPSTAble As ObjectId = Nothing
            If boolPSBuild = True Then
                IDPSTAble = FuncPropertySet.FuncReturnIDDictSetProperties(NamePSSetTable)
            End If
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбор файла
                Dim NameFiles As String() = {""}
                Dim fileDialog1 As OpenFileDialog = New OpenFileDialog("Открыть файл xml", "", "xml", "xml", OpenFileDialogFlags.AllowMultiple)
                Dim FileRez As String = fileDialog1.ShowDialog()
                NameFiles = fileDialog1.GetFilenames
                If IsNothing(NameFiles) = True Then Exit Sub
                Dim arrayCoord As String(,) = {}
                Dim countArrayDouble As Integer = 0
                Dim boolNode As Boolean = False
                For i As Integer = 0 To NameFiles.GetUpperBound(0)
                    Dim FileName As String = NameFiles(i)
                    Dim nodeDictionary As Dictionary(Of String, Double()) = New Dictionary(Of String, Double())
                    Using reader As XmlReader = XmlReader.Create(FileName)
                        While reader.Read()
                            If reader.NodeType = XmlNodeType.Element AndAlso reader.Name = "node" Then
                                Dim id As String = reader.GetAttribute("id")
                                Dim lat As String = reader.GetAttribute("lat")
                                Dim lon As String = reader.GetAttribute("lon")
                                Dim arrayDictCoord As Double() = {CDbl(lat), CDbl(lon)}
                                nodeDictionary.Add(id, arrayDictCoord)
                                ' Чтение вложенных тегов (если есть)
                                'If reader.ReadToDescendant("tag") Then
                                '    Do
                                '        Dim key As String = reader.GetAttribute("k")
                                '        Dim value As String = reader.GetAttribute("v")
                                '        Console.WriteLine($"  Tag: {key} = {value}")
                                '    Loop While reader.ReadToNextSibling("tag")
                                'End If
                            End If
                        End While
                    End Using

                    Dim pLineDict As Dictionary(Of String, Polyline) = New Dictionary(Of String, Polyline)
                    Using reader1 As XmlReader = XmlReader.Create(FileName)
                        While reader1.Read()
                            If reader1.NodeType = XmlNodeType.Element AndAlso reader1.Name = "way" Then
                                Dim wayId As String = reader1.GetAttribute("id")
                                ' Чтение ссылок на узлы (nd)
                                Dim IDPLine As ObjectId = Nothing
                                If reader1.ReadToFollowing("nd") Then
                                    Dim newCoordArray As String(,) = {}
                                    Dim countArray As Integer = 0
                                    Dim LL84WGS_to_DrawCoord As MgCoordinateSystemTransform = coordSysFactory.GetTransform(LL84Cs, DrawingSysCoord)
                                    Do
                                        Dim refId As String = reader1.GetAttribute("ref")
                                        If nodeDictionary.ContainsKey(refId) = True Then
                                            Dim arrayData As Double() = nodeDictionary.Item(refId)
                                            Dim lat As Double = arrayData(0)
                                            Dim lon As Double = arrayData(1)
                                            Dim LL84WGS_to_DrawCoordRez As MgCoordinate = LL84WGS_to_DrawCoord.Transform(lon, lat)
                                            Dim xrez As Double = LL84WGS_to_DrawCoordRez.X
                                            Dim yrez As Double = LL84WGS_to_DrawCoordRez.Y
                                            ReDim Preserve newCoordArray(3, countArray)
                                            newCoordArray(3, countArray) = xrez
                                            newCoordArray(2, countArray) = yrez
                                            countArray += 1
                                        End If
                                    Loop While reader1.ReadToNextSibling("nd")
                                    'если число вершин больше 1 тогда рисуем полилинию
                                    If newCoordArray.Length > 0 Then
                                        If newCoordArray.GetUpperBound(1) > 0 Then
                                            IDPLine = FuncCadastr.FuncDrawPolyline(newCoordArray)
                                            If IsNothing(IDPLine) = False Then
                                                If IDPLine.IsNull = False Then
                                                    Dim ArrayData As String(,) = Nothing
                                                    ReDim Preserve ArrayData(2, 15)
                                                    ArrayData(0, 0) = NamePSSetTable
                                                    ArrayData(1, 0) = "Cad_number"
                                                    ArrayData(2, 0) = ""

                                                    ArrayData(0, 1) = NamePSSetTable
                                                    ArrayData(1, 1) = "Short_cad_number"
                                                    ArrayData(2, 1) = ""

                                                    ArrayData(0, 2) = NamePSSetTable
                                                    ArrayData(1, 2) = "State"
                                                    ArrayData(2, 2) = ""

                                                    ArrayData(0, 3) = NamePSSetTable
                                                    ArrayData(1, 3) = "Assignation"
                                                    ArrayData(2, 3) = ""

                                                    ArrayData(0, 4) = NamePSSetTable
                                                    ArrayData(1, 4) = "MaterialWall"
                                                    ArrayData(2, 4) = ""

                                                    ArrayData(0, 5) = NamePSSetTable
                                                    ArrayData(1, 5) = "FloatCount"
                                                    ArrayData(2, 5) = ""

                                                    ArrayData(0, 6) = NamePSSetTable
                                                    ArrayData(1, 6) = "Area"
                                                    ArrayData(2, 6) = ""

                                                    ArrayData(0, 7) = NamePSSetTable
                                                    ArrayData(1, 7) = "OKATO"
                                                    ArrayData(2, 7) = ""

                                                    ArrayData(0, 8) = NamePSSetTable
                                                    ArrayData(1, 8) = "Adress"
                                                    ArrayData(2, 8) = ""

                                                    ArrayData(0, 9) = NamePSSetTable
                                                    ArrayData(1, 9) = "Right_type"
                                                    ArrayData(2, 9) = ""

                                                    ArrayData(0, 10) = NamePSSetTable
                                                    ArrayData(1, 10) = "Patronymic"
                                                    ArrayData(2, 10) = ""

                                                    ArrayData(0, 11) = NamePSSetTable
                                                    ArrayData(1, 11) = "Encumbrance"
                                                    ArrayData(2, 11) = ""

                                                    ArrayData(0, 12) = NamePSSetTable
                                                    ArrayData(1, 12) = "CadastralCost"
                                                    ArrayData(2, 12) = ""

                                                    ArrayData(0, 13) = NamePSSetTable
                                                    ArrayData(1, 13) = "DateCreated"
                                                    ArrayData(2, 13) = ""

                                                    ArrayData(0, 14) = NamePSSetTable
                                                    ArrayData(1, 14) = "LOCALID"
                                                    ArrayData(2, 14) = wayId

                                                    ArrayData(0, 15) = NamePSSetTable
                                                    ArrayData(1, 15) = "GlobalID"
                                                    ArrayData(2, 15) = ""

                                                    Dim pline As Polyline = IDPLine.GetObject(OpenMode.ForWrite)
                                                    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(pline, IDPSTAble)
                                                    If boolPSObject = True Then
                                                        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(pline, ArrayData)
                                                    End If
                                                    pLineDict.Add(wayId, pline)
                                                    'End If
                                                    'Dim pline As Polyline = IDPLine.GetObject(OpenMode.ForWrite)
                                                    'Dim street As String = ""
                                                    'Dim hours As String = ""
                                                    'Dim post As String = ""
                                                    'Dim boolBuild As Boolean = False
                                                    'Do
                                                    '    Dim key As String = reader1.GetAttribute("k")
                                                    '    Dim value As String = reader1.GetAttribute("v")
                                                    '    If key Like "building" Then
                                                    '        pline.ColorIndex = 10
                                                    '        boolBuild = True
                                                    '    ElseIf key Like "building:levels" Then
                                                    '        ArrayData(2, 5) = value
                                                    '    ElseIf key Like "addr:street" Then
                                                    '        street = value
                                                    '    ElseIf key Like "addr:postcode" Then
                                                    '        post = value
                                                    '    ElseIf key Like "addr:housenumber" Then
                                                    '        hours = value
                                                    '    End If
                                                    'Loop While reader1.ReadToNextSibling("tag")
                                                    'If boolBuild = True Then
                                                    '    ArrayData(2, 8) = post & " " & street & " " & hours
                                                    '    Dim boolPSObject As Boolean = FuncPropertySet.FuncAddDataSetProperties(pline, IDPSTAble)
                                                    '    If boolPSObject = True Then
                                                    '        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(pline, ArrayData)
                                                    '    End If
                                                    'End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End While
                    End Using

                    Using reader2 As XmlReader = XmlReader.Create(FileName)
                        While reader2.Read()
                            If reader2.NodeType = XmlNodeType.Element AndAlso reader2.Name = "way" Then
                                Dim wayId As String = reader2.GetAttribute("id")
                                If reader2.ReadToFollowing("tag") Then
                                    Dim street As String = ""
                                    Dim hours As String = ""
                                    Dim post As String = ""
                                    Dim numberFloo As String = ""
                                    Dim boolBuild As Boolean = False
                                    Do
                                        Dim key As String = reader2.GetAttribute("k")
                                        Dim value As String = reader2.GetAttribute("v")
                                        If key Like "building" Then
                                            boolBuild = True
                                        ElseIf key Like "building:levels" Then
                                            numberFloo = value
                                        ElseIf key Like "addr:street" Then
                                            street = value
                                        ElseIf key Like "addr:postcode" Then
                                            post = value
                                        ElseIf key Like "addr:housenumber" Then
                                            hours = value
                                        End If
                                    Loop While reader2.ReadToNextSibling("tag")
                                    If boolBuild = True Then
                                        Dim pline As Polyline = pLineDict.Item(wayId)
                                        pline.ColorIndex = 10
                                        Dim ArrayData As String(,) = Nothing
                                        ReDim Preserve ArrayData(2, 2)

                                        ArrayData(0, 0) = NamePSSetTable
                                        ArrayData(1, 0) = "FloatCount"
                                        ArrayData(2, 0) = numberFloo

                                        ArrayData(0, 1) = NamePSSetTable
                                        ArrayData(1, 1) = "Area"
                                        ArrayData(2, 1) = Math.Round(pline.Area, 2)

                                        ArrayData(0, 2) = NamePSSetTable
                                        ArrayData(1, 2) = "Adress"
                                        ArrayData(2, 2) = post & " " & street & " " & hours
                                        Dim boolPSData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(pline, ArrayData)
                                    End If
                                End If
                            End If
                        End While
                    End Using
                Next i
                acTrans.Commit()
            End Using
        End Using

    End Sub



    '=============================================================================================
    'Создать сетку Face на основе выбранной полилинии
    '=============================================================================================
    <CommandMethod("CreateFaceByPolyline")>
    Public Sub CreateFaceByPolyline()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите полилинию: ")
            If Not (acSSPr.Status = PromptStatus.OK) Then
                Exit Sub
            End If

            Dim entPoly As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
            If Not (TypeOf entPoly Is Polyline) Then
                MsgBox("Выбранный примитив не является Полилинией. Макро прервано.")
                Exit Sub
            End If
            Dim pPoly As Polyline = entPoly
            'массив с таблицами PS
            Dim ArrayTablesPS As String() = Nothing
            Dim boolTablePs As Boolean = FuncPropertySet.FuncReadAllDataSetTables(pPoly, ArrayTablesPS)
            'массив с пользовательскими свойствами
            Dim ArrayData As String(,) = Nothing
            Dim boolPS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(pPoly, ArrayData)
            '==============================================================================================================================
            'создаем новую сеть
            Dim extrSurf As Autodesk.AutoCAD.DatabaseServices.ExtrudedSurface = New Autodesk.AutoCAD.DatabaseServices.ExtrudedSurface()
            Dim Face As Autodesk.AutoCAD.DatabaseServices.Face = Nothing
            If pPoly.StartPoint = pPoly.EndPoint Then
                Face = New Autodesk.AutoCAD.DatabaseServices.Face(pPoly.StartPoint, pPoly.GetPoint3dAt(pPoly.NumberOfVertices - 1), pPoly.GetPoint3dAt(pPoly.NumberOfVertices - 1) + New Vector3d(0, 0, 1000), pPoly.StartPoint + New Vector3d(0, 0, 1000), True, True, True, True)
            Else
                Face = New Autodesk.AutoCAD.DatabaseServices.Face(pPoly.StartPoint, pPoly.EndPoint, pPoly.EndPoint + New Vector3d(0, 0, 1000), pPoly.StartPoint + New Vector3d(0, 0, 1000), True, True, True, True)
            End If
            Dim sweepEnt As Autodesk.AutoCAD.DatabaseServices.Surface = New Autodesk.AutoCAD.DatabaseServices.Surface()

            sweepEnt.SetDatabaseDefaults()
            Try
                sweepEnt = Autodesk.AutoCAD.DatabaseServices.Surface.CreateFrom(Face)
                Dim sweepOpts As SweepOptions = New SweepOptions()
                extrSurf.CreateExtrudedSurface(sweepEnt, (pPoly.EndPoint - pPoly.StartPoint).GetPerpendicularVector(), sweepOpts)
                If IsArray(ArrayData) = True And IsArray(ArrayTablesPS) Then
                    'присваиваем таблицы
                    For i As Integer = 0 To ArrayTablesPS.Length - 1
                        Dim IDTbl As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTablesPS(i).Trim)
                        If IsNothing(IDTbl) = False Then
                            If IDTbl.IsNull = False Then
                                Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(extrSurf, IDTbl)
                            End If
                        End If
                    Next i
                    Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(extrSurf, ArrayData)
                End If
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                Ed.WriteMessage("\nFailed with CreateExtrudedSurface.")
            End Try
            acTrans.Commit()
        End Using
    End Sub
    '=============================================================================================
    'Создать 3d тела из выбранных полилиний на основе исходного 3d Solid (исходное тело не корректируется)
    '=============================================================================================
    <CommandMethod("CreateSolid3dByPolyline")>
    Public Sub CreateSolid3dByPolyline()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        'выбираем тело для обрезки
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите Solid3D: ")
            If Not (acSSPr.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim entSurfaceSolid As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
            If Not (TypeOf entSurfaceSolid Is Solid3d) Then
                MsgBox("Выбранный примитив не является солидом. Макро прервано.")
                Exit Sub
            End If
            Dim SurfaceSolid3d As Solid3d = entSurfaceSolid
            'находим максимум и минимум по высоте для формирования вспомогательного солида для обрезки
            Dim Zmin As Double = SurfaceSolid3d.GeometricExtents.MinPoint.Z
            Dim Zmax As Double = SurfaceSolid3d.GeometricExtents.MaxPoint.Z
            If Zmin = Zmax Then
                MsgBox("Солид не имеет толщины. Макро прервано!")
                Exit Sub
            End If
            'выбираем полилинию
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Polyline Then
                                    Dim Poly1 As Polyline = acEnt
                                    Poly1.Elevation = Zmax
                                    'массив с таблицами PS
                                    Dim ArrayTablesPS As String() = Nothing
                                    Dim boolTablePs As Boolean = FuncPropertySet.FuncReadAllDataSetTables(Poly1, ArrayTablesPS)
                                    'массив с пользовательскими свойствами
                                    Dim ArrayData As String(,) = Nothing
                                    Dim boolPS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(Poly1, ArrayData)
                                    Dim Poly2 As Polyline = Poly1.Clone()
                                    Poly2.Elevation = Zmin
                                    Dim options As New LoftOptions
                                    Dim sol3d As New Solid3d
                                    sol3d.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {Poly1, Poly2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, Nothing, options)
                                    acBlkTblRec.AppendEntity(sol3d)
                                    acTrans.AddNewlyCreatedDBObject(sol3d, True)
                                    Dim acSol3DCopy3 As Solid3d = SurfaceSolid3d.Clone()
                                    acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, sol3d.Clone())
                                    acBlkTblRec.AppendEntity(acSol3DCopy3)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)
                                    acSol3DCopy3.Layer = Poly1.Layer
                                    sol3d.Erase()
                                    If IsArray(ArrayData) = True And IsArray(ArrayTablesPS) Then
                                        'присваиваем таблицы
                                        For i As Integer = 0 To ArrayTablesPS.Length - 1
                                            Dim IDTbl As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTablesPS(i).Trim)
                                            If IsNothing(IDTbl) = False Then
                                                If IDTbl.IsNull = False Then
                                                    Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(acSol3DCopy3, IDTbl)
                                                End If
                                            End If
                                        Next i
                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSol3DCopy3, ArrayData)
                                    End If
                                End If
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
                acTrans.Commit()
            End If
        End Using
    End Sub
    '=============================================================================================
    'Обрезка solid3d по полилинии (исходное тело корректируется, т.е. появляется дырка)
    '=============================================================================================
    <CommandMethod("TrimSolid3dByPolyline")>
    Public Sub TrimSolid3dByPolyline()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        'выбираем тело для обрезки
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите полилинию по которой будет производится обрезка Solid3d: ")
            If Not (acSSPr.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim entpLineBoundares As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
            If Not (TypeOf entpLineBoundares Is Polyline) Then
                MsgBox("Выбранный примитив не является полилинией. Макро прервано.")
                Exit Sub
            End If
            Dim pLineBoundares As Polyline = entpLineBoundares
            'находим максимум и минимум по высоте для формирования вспомогательного солида для обрезки
            'выбираем солиды для обрезки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Solid3d Then
                                    Dim SurfaceSolid3d As Solid3d = acEnt
                                    Dim Zmin As Double = SurfaceSolid3d.GeometricExtents.MinPoint.Z
                                    Dim Zmax As Double = SurfaceSolid3d.GeometricExtents.MaxPoint.Z
                                    If Zmin = Zmax Then
                                        Continue For
                                    Else
                                        pLineBoundares.Elevation = Zmax
                                    End If
                                    'массив с таблицами PS
                                    Dim ArrayTablesPS As String() = Nothing
                                    Dim boolTablePs As Boolean = FuncPropertySet.FuncReadAllDataSetTables(SurfaceSolid3d, ArrayTablesPS)
                                    'массив с пользовательскими свойствами
                                    Dim ArrayData As String(,) = Nothing
                                    Dim boolPS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(SurfaceSolid3d, ArrayData)

                                    Dim Poly2 As Polyline = pLineBoundares.Clone()
                                    Poly2.Elevation = Zmin
                                    Dim options As New LoftOptions
                                    Dim sol3d As New Solid3d
                                    sol3d.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {pLineBoundares, Poly2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, Nothing, options)
                                    acBlkTblRec.AppendEntity(sol3d)
                                    acTrans.AddNewlyCreatedDBObject(sol3d, True)

                                    Dim acSol3DCopy3 As Solid3d = SurfaceSolid3d.Clone()
                                    acBlkTblRec.AppendEntity(acSol3DCopy3)
                                    acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)

                                    acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, sol3d.Clone())
                                    Dim UserLayerId As ObjectId = acCurDb.Clayer
                                    acSol3DCopy3.LayerId = SurfaceSolid3d.LayerId

                                    SurfaceSolid3d.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DCopy3.Clone())

                                    If IsArray(ArrayData) = True And IsArray(ArrayTablesPS) Then
                                        'присваиваем таблицы
                                        For i As Integer = 0 To ArrayTablesPS.Length - 1
                                            Dim IDTbl As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTablesPS(i).Trim)
                                            If IsNothing(IDTbl) = False Then
                                                If IDTbl.IsNull = False Then
                                                    Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(SurfaceSolid3d, IDTbl)
                                                End If
                                            End If
                                        Next i
                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(SurfaceSolid3d, ArrayData)
                                    End If
                                    sol3d.Erase()
                                End If
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
                acTrans.Commit()
            End If
        End Using
    End Sub
    '=============================================================================================
    'Обрезка выбранного солида, другим солидом
    '=============================================================================================
    <CommandMethod("ObrezkaUserSolid")>
    Public Sub ObrezkaSolid()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Do
            Try
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    'выбираем тело по которому будет происходит обрезка
                    Dim acSSPrompt1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело по которому будет происходить обрезка: ")
                    If Not (acSSPrompt1.Status = PromptStatus.OK) Then
                        Exit Do
                    End If
                    Dim TrimSolid3d As Solid3d = acTrans.GetObject(acSSPrompt1.ObjectId, OpenMode.ForRead)
                    'выбираем тела для обрезки
                    Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim acSSet As SelectionSet = acSSPrompt.Value
                        For Each acSSObj As SelectedObject In acSSet
                            If Not IsDBNull(acSSObj) Then
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                If acEnt.Id = TrimSolid3d.Id Then Continue For
                                If Not IsDBNull(acEnt) Then
                                    If TypeOf acEnt Is Solid3d Then
                                        Dim acSolid3d As Solid3d = acEnt
                                        If acSolid3d.CheckInterference(TrimSolid3d) = True Then
                                            Dim ArrayTables As String() = Nothing
                                            Dim boolReadPSTable As Boolean = FuncPropertySet.FuncReadAllDataSetTables(acSolid3d, ArrayTables)
                                            Dim ArrayData As String(,) = Nothing
                                            If IsArray(ArrayTables) = True Then
                                                Dim boolReadDataPs As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acSolid3d, ArrayData)
                                            End If
                                            Dim acSol3DCopy As Solid3d = acSolid3d.Clone()
                                            acSol3DCopy.BooleanOperation(BooleanOperationType.BoolSubtract, TrimSolid3d.Clone())
                                            acBlkTblRec.AppendEntity(acSol3DCopy)
                                            acTrans.AddNewlyCreatedDBObject(acSol3DCopy, True)
                                            If IsArray(ArrayData) = True And IsArray(ArrayTables) = True Then
                                                For i As Integer = 0 To ArrayTables.Length - 1
                                                    Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTables(i))
                                                    If IsNothing(CreateTBLPSID) = False Then
                                                        If CreateTBLPSID.IsNull = False Then
                                                            Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acSol3DCopy, CreateTBLPSID)
                                                        End If
                                                    End If
                                                Next i
                                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSol3DCopy, ArrayData)
                                            End If
                                            acSolid3d.Erase()
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                    acTrans.Commit()
                End Using
            Catch ex As System.Exception
            End Try
        Loop
    End Sub
    '=============================================================================================
    'Обрезка всех солидов с анализом пересечений
    '=============================================================================================
    <CommandMethod("ObrezkaUserAllSolid")>
    Public Sub ObrezkaUserAllSolid()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim ArrayUserSolid As String(,) = Nothing
                Dim countArrayUserSolid As Integer = 0
                'выбираем тела для обрезки
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If TypeOf acEnt Is Solid3d Then
                                Dim TempSolid As Solid3d = acEnt
                                ReDim Preserve ArrayUserSolid(5, countArrayUserSolid)
                                ArrayUserSolid(0, countArrayUserSolid) = TempSolid.MassProperties.Volume
                                ArrayUserSolid(1, countArrayUserSolid) = TempSolid.GeometricExtents.MinPoint.X
                                ArrayUserSolid(2, countArrayUserSolid) = TempSolid.GeometricExtents.MinPoint.Y
                                ArrayUserSolid(3, countArrayUserSolid) = TempSolid.GeometricExtents.MaxPoint.X
                                ArrayUserSolid(4, countArrayUserSolid) = TempSolid.GeometricExtents.MaxPoint.Y
                                ArrayUserSolid(5, countArrayUserSolid) = TempSolid.Handle.ToString
                                countArrayUserSolid += 1
                            End If
                        End If
                    Next
                End If
                If IsArray(ArrayUserSolid) = False Then
                    MsgBox("В выбранном наборе нет объектов типа Solid3d!!! Макро преревано.")
                    Exit Sub
                End If
                If ArrayUserSolid.GetUpperBound(1) < 1 Then
                    MsgBox("В выбранном наборе только один объект типа Solid3d!!! Макро преревано.")
                    Exit Sub
                End If
                'сортируем массив по объему солидов
                Dim boolSortSolid As Boolean = MathFunction.FuncSortStrArray(ArrayUserSolid, 0)
                'достаем их последовательно
                For i As Integer = 0 To ArrayUserSolid.GetUpperBound(1) - 1
                    Dim volumeSolid As Double = Val(ArrayUserSolid(0, i))
                    Dim minXSolid As Double = Val(ArrayUserSolid(1, i))
                    Dim minYSolid As Double = Val(ArrayUserSolid(2, i))
                    Dim maxXSolid As Double = Val(ArrayUserSolid(3, i))
                    Dim maxYSolid As Double = Val(ArrayUserSolid(4, i))
                    Dim idSolid3d As ObjectId = AcadFunction.FuncReturnObjectID(ArrayUserSolid(5, i))
                    If IsNothing(idSolid3d) = False Then
                        If idSolid3d.IsNull = False Then
                            Dim UserSolid3d As Solid3d = acTrans.GetObject(idSolid3d, OpenMode.ForWrite)
                            'флаг указывающий на то что пересечение найдено
                            Dim boolErase As Boolean = False
                            For j As Integer = 1 To ArrayUserSolid.GetUpperBound(1)
                                Dim idNextSolid3d As ObjectId = AcadFunction.FuncReturnObjectID(ArrayUserSolid(5, j))
                                If IsNothing(idNextSolid3d) = False Then
                                    If idNextSolid3d.IsNull = False Then
                                        Dim UserNextSolid3d As Solid3d = acTrans.GetObject(idNextSolid3d, OpenMode.ForWrite)
                                        If UserNextSolid3d.CheckInterference(UserSolid3d) = True Then
                                            UserNextSolid3d.BooleanOperation(BooleanOperationType.BoolSubtract, UserSolid3d.Clone())
                                            boolErase = True
                                        End If
                                    End If
                                End If
                            Next j
                        End If
                    End If
                Next i
                acTrans.Commit()
            End Using
        Catch ex As System.Exception
        End Try
        'If boolErase = True Then
        '    If IsArray(ArrayData) = True And IsArray(ArrayTables) = True Then
        '        For k As Integer = 0 To ArrayTables.Length - 1
        '            Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTables(k))
        '            If IsNothing(CreateTBLPSID) = False Then
        '                If CreateTBLPSID.IsNull = False Then
        '                    Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(UserSolid3d, CreateTBLPSID)
        '                End If
        '            End If
        '        Next k
        '        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(UserSolid3d, ArrayData)
        '    End If
        'End If
        'Dim ArrayTables As String() = Nothing
        'Dim boolReadPSTable As Boolean = FuncPropertySet.FuncReadAllDataSetTables(UserSolid3d, ArrayTables)
        'Dim ArrayData As String(,) = Nothing
        'If IsArray(ArrayTables) = True Then
        '    Dim boolReadDataPs As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(UserSolid3d, ArrayData)
        'End If
        'Dim minXTempSolid As Double = ArrayUserSolid(1, j)
        'Dim minYTempSolid As Double = ArrayUserSolid(2, j)
        'Dim maxXTempSolid As Double = ArrayUserSolid(3, j)
        ' Dim maxYTempSolid As Double = ArrayUserSolid(4, j)
        'Dim boolIntersect As Boolean = MathFunction.FuncIntersectedRectangle(minXSolid, maxYSolid, maxXSolid, minYSolid, minXTempSolid, maxYTempSolid, maxXTempSolid, minYTempSolid)

        ' Dim UserTempSolid3d As Solid3d = acTrans.GetObject(idTempSolid3d, OpenMode.ForWrite)

    End Sub
    '=============================================================================================
    'создать 3d Solid по 3d полилинии
    <CommandMethod("CreateSolid3dByPolyline3d")>
    Public Sub CreateSolid3dByPolyline3d()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите 3d полилинию по которой будет производится создание Solid3d: ")
                If Not (acSSPr.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entpLineBoundares As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
                If Not (TypeOf entpLineBoundares Is Polyline3d) Then
                    MsgBox("Выбранный примитив не является 3d полилинией. Макро прервано.")
                    Exit Sub
                End If
                Dim pLineBoundares As Polyline3d = entpLineBoundares
                'удаляем лишние вершины
                Dim boolDeleteDoubleVertex As Boolean = AcadFunction.FuncPlineDelDubleVertex(pLineBoundares)
                'выбираем солиды для обрезки
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        Try
                            If Not IsDBNull(acSSObj) Then
                                Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                If Not IsDBNull(acEnt) Then
                                    If TypeOf acEnt Is Solid3d Then
                                        Dim SurfaceSolid3d As Solid3d = acEnt
                                        Dim Zmin As Double = SurfaceSolid3d.GeometricExtents.MinPoint.Z
                                        Dim Zmax As Double = SurfaceSolid3d.GeometricExtents.MaxPoint.Z
                                        If Zmin = Zmax Then
                                            Continue For
                                        Else
                                            'pLineBoundares.Elevation = Zmax
                                        End If
                                        'массив с таблицами PS
                                        Dim ArrayTablesPS As String() = Nothing
                                        Dim boolTablePs As Boolean = FuncPropertySet.FuncReadAllDataSetTables(SurfaceSolid3d, ArrayTablesPS)
                                        'массив с пользовательскими свойствами
                                        Dim ArrayData As String(,) = Nothing
                                        Dim boolPS As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(SurfaceSolid3d, ArrayData)

                                        Dim Poly2 As Polyline = pLineBoundares.Clone()
                                        Poly2.Elevation = Zmin
                                        Dim options As New LoftOptions
                                        Dim sol3d As New Solid3d
                                        sol3d.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {pLineBoundares, Poly2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, Nothing, options)
                                        acBlkTblRec.AppendEntity(sol3d)
                                        acTrans.AddNewlyCreatedDBObject(sol3d, True)

                                        Dim acSol3DCopy3 As Solid3d = SurfaceSolid3d.Clone()
                                        acBlkTblRec.AppendEntity(acSol3DCopy3)
                                        acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)

                                        acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, sol3d.Clone())
                                        Dim UserLayerId As ObjectId = acCurDb.Clayer
                                        acSol3DCopy3.LayerId = UserLayerId

                                        SurfaceSolid3d.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DCopy3.Clone())

                                        If IsArray(ArrayData) = True And IsArray(ArrayTablesPS) Then
                                            'присваиваем таблицы
                                            For i As Integer = 0 To ArrayTablesPS.Length - 1
                                                Dim IDTbl As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayTablesPS(i).Trim)
                                                If IsNothing(IDTbl) = False Then
                                                    If IDTbl.IsNull = False Then
                                                        Dim boolAddTable As Boolean = FuncPropertySet.FuncAddDataSetProperties(SurfaceSolid3d, IDTbl)
                                                    End If
                                                End If
                                            Next i
                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(SurfaceSolid3d, ArrayData)
                                        End If
                                        sol3d.Erase()
                                    End If
                                End If
                            End If
                        Catch ex As System.Exception
                        End Try
                    Next
                    acTrans.Commit()
                End If
            End Using

        Catch ex As System.Exception
        End Try
    End Sub
    '=============================================================================================
    'фаска на полилинии
    Public Shared Sub FilletAll(ByVal pline As Polyline, ByVal radius As Double)
        Dim n As Integer = If(pline.Closed, 0, 1)
        Dim i As Integer = n

        While i < pline.NumberOfVertices - n
            i += 1 + FilletAt(pline, i, radius)
        End While
    End Sub

    <CommandMethod("filletUserPline")>
    Public Shared Sub filletUserPline()
        Dim doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db = doc.Database
        Dim ed = doc.Editor
        Dim filter = New SelectionFilter({New TypedValue(0, "LWPOLYLINE")})
        Dim selection = ed.GetSelection(filter)
        If selection.Status <> PromptStatus.OK Then Return
        Dim pdo = New PromptDistanceOptions(vbLf & "Введите радиус кривой: ")
        pdo.DefaultValue = CDbl(Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("FILLETRAD"))
        Dim pdr = ed.GetDistance(pdo)
        If pdr.Status <> PromptStatus.OK Then Return
        Autodesk.AutoCAD.ApplicationServices.Core.Application.SetSystemVariable("FILLETRAD", pdr.Value)
        Using tr = db.TransactionManager.StartTransaction()
            For Each selectedObject As SelectedObject In selection.Value
                Dim pline = CType(tr.GetObject(selectedObject.ObjectId, OpenMode.ForWrite), Polyline)
                IFC.FilletAll(pline, pdr.Value)
            Next
            tr.Commit()
        End Using
    End Sub








    '=============================================================================================
    'Создание водопропускной трубы
    ''=============================================================================================
    <CommandMethod("CreatePipeFillSolid")>
    Public Sub CreatePipeFillSolid()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            Dim peo1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите характерную линию: ")
            Dim per As PromptEntityResult = ed.GetEntity(peo1)
            If per.Status = PromptStatus.OK Then
                Dim regId As ObjectId = per.ObjectId
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(regId, OpenMode.ForRead)
                If Not (TypeOf ent Is Autodesk.Civil.DatabaseServices.FeatureLine) Then
                    MsgBox("Ошибка выбора объектов. Необходимо выбрать Характерную линию!!!")
                    Exit Sub
                End If
                'назначаем слой
                Dim acFeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = ent
                Dim nameStyleFeatLine As String = acFeatLine.StyleName
                nameStyleFeatLine = nameStyleFeatLine.Replace("ИИ-", "")
                Dim nameLayer As String = "IFC-" & nameStyleFeatLine
                Dim boolAddlayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer)
                'читаем диаметр трубы
                Dim strDiameter As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "DIAMETER")
                Dim ValDiameter As Double = Val(strDiameter)
                Dim ValHeight As Double = 0
                Dim ValWidth As Double = 0
                If ValDiameter = 0 Then
                    Dim strHeight As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "HEIGHT")
                    ValHeight = Val(strHeight) / 1000
                    Dim strWight As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "WIDTH")
                    ValWidth = Val(strWight) / 1000
                Else
                    ValDiameter = ValDiameter / 1000
                End If
                'читаем толщину стенки
                Dim WallT As Double = 0
                Dim strVallT As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "WALL_THICKNESS")
                If Val(strVallT) > 0 Then
                    WallT = Val(strVallT)
                Else
                    Dim optDb1 As PromptDoubleOptions = New PromptDoubleOptions("Введите толщину стенки трубы: ")
                    Dim resDb1 As PromptDoubleResult = ed.GetDouble(optDb1)
                    WallT = Math.Abs(resDb1.Value)
                End If
                WallT = WallT / 1000
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim acCirc1 As Circle = New Circle()
                'читаем все данные
                Dim ArrayData As String(,) = Nothing
                Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acFeatLine, ArrayData)
                Dim IDTablePS As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayData(0, 0))
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                Dim PointColl As Point3dCollection = acFeatLine.GetPoints(ftype)
                If PointColl.Count > 1 Then
                    Using acPoly3d As Polyline3d = New Polyline3d()
                        acBlkTblRec.AppendEntity(acPoly3d)
                        tr.AddNewlyCreatedDBObject(acPoly3d, True)
                        For Each pt As Point3d In PointColl
                            If ValDiameter > 0 Then
                                Dim newPt As Point3d = New Point3d(pt.X, pt.Y, pt.Z - ValDiameter / 2)
                                Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(newPt)
                                    acPoly3d.AppendVertex(acPolVer3d)
                                    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                                End Using
                            ElseIf ValHeight > 0 Then
                                Dim newPt As Point3d = New Point3d(pt.X, pt.Y, pt.Z - ValHeight / 2)
                                Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(newPt)
                                    acPoly3d.AppendVertex(acPolVer3d)
                                    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                                End Using
                            End If
                        Next
                        'рисуем окружность 1 с внутренним диаметром
                        If ValDiameter > 0 Then
                            If WallT > 0 Then
                                acBlkTblRec.AppendEntity(acCirc1)
                                tr.AddNewlyCreatedDBObject(acCirc1, True)
                                acCirc1.Diameter = ValDiameter + WallT
                                acCirc1.Center = acPoly3d.StartPoint
                            End If
                        Else
                            MsgBox("YНет данных по диаметру трубы!!! Макро прервано")
                            Exit Sub
                        End If

                        Dim sob As SweepOptionsBuilder = New SweepOptionsBuilder()
                        sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
                        sob.BasePoint = acPoly3d.StartPoint
                        sob.Bank = True

                        'создаем 2 солида
                        Dim sol1 As Solid3d = New Solid3d()
                        Dim sol2 As Solid3d = New Solid3d()
                        If ValDiameter > 0 Then
                            sol1.CreateSweptSolid(acCirc1, acPoly3d, sob.ToSweepOptions())
                            acBlkTblRec.AppendEntity(sol1)
                            tr.AddNewlyCreatedDBObject(sol1, True)
                        End If
                        acCirc1.Erase()
                        acPoly3d.Erase()
                    End Using
                End If
                tr.Commit()
            End If
        End Using
    End Sub




    '=============================================================================================
    'Поставить многовидовые блоки на поверхность
    '=============================================================================================
    <CommandMethod("RefineElevationsMVBlock")>
    Public Sub RefineElevationsMVBlock()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        'выбираем тело для обрезки
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim userSurface As TinSurface = Nothing
            If CivilDoc.GetSurfaceIds.Count = 1 Then
                Dim IDSurf As ObjectId = CivilDoc.GetSurfaceIds.Item(0)
                If IDSurf.IsNull = False Then
                    userSurface = acTrans.GetObject(IDSurf, OpenMode.ForRead)
                End If
            ElseIf CivilDoc.GetSurfaceIds.Count > 1 Then
                Dim acSSPr As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите Поверхность TIN: ")
                If Not (acSSPr.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entSurface As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPr.ObjectId, OpenMode.ForWrite)
                If Not (TypeOf entSurface Is TinSurface) Then
                    MsgBox("Выбранный примитив не является Поверхностью TIN. Макро прервано.")
                    Exit Sub
                End If
                userSurface = entSurface
            Else
                MsgBox("В чертеже нет поверхности TIN. Макро прервано!!!")
                Exit Sub
            End If

            'выбираем многовидовые блоки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    Try
                        If Not IsDBNull(acSSObj) Then
                            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeOf acEnt Is Autodesk.AEC.DatabaseServices.MultiViewBlockReference Then
                                    Dim mvb As Autodesk.AEC.DatabaseServices.MultiViewBlockReference = acEnt
                                    Dim PositionMVB As Point3d = mvb.Location
                                    Try
                                        Dim elev As Double = userSurface.FindElevationAtXY(PositionMVB.X, PositionMVB.Y)
                                        Dim InsPoint As Point3d = New Point3d(PositionMVB.X, PositionMVB.Y, elev)
                                        mvb.Location = InsPoint
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    End Try
                                End If
                            End If
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
                acTrans.Commit()
            End If
        End Using
    End Sub






    '=============================================================================================
    'Создание водопропускной трубы
    ''=============================================================================================
    <CommandMethod("CreatePipeSolid")>
    Public Sub CreatePipeSolid()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
        Using tr
            Dim peo1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите характерную линию: ")
            Dim per As PromptEntityResult = ed.GetEntity(peo1)
            If per.Status = PromptStatus.OK Then
                Dim regId As ObjectId = per.ObjectId
                Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(regId, OpenMode.ForRead)
                If Not (TypeOf ent Is Autodesk.Civil.DatabaseServices.FeatureLine) Then
                    MsgBox("Ошибка выбора объектов. Необходимо выбрать Характерную линию!!!")
                    Exit Sub
                End If
                'назначаем слой
                Dim acFeatLine As Autodesk.Civil.DatabaseServices.FeatureLine = ent
                Dim nameStyleFeatLine As String = acFeatLine.StyleName
                nameStyleFeatLine = nameStyleFeatLine.Replace("ИИ-", "")
                Dim nameLayer As String = "IFC-" & nameStyleFeatLine
                Dim boolAddlayer As ObjectId = AcadFunction.FuncAddLayer(nameLayer)
                'читаем диаметр трубы
                Dim strDiameter As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "DIAMETER")
                Dim ValDiameter As Double = Val(strDiameter)
                Dim ValHeight As Double = 0
                Dim ValWidth As Double = 0
                If ValDiameter = 0 Then
                    Dim strHeight As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "HEIGHT")
                    ValHeight = Val(strHeight) / 1000
                    Dim strWight As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "WIDTH")
                    ValWidth = Val(strWight) / 1000
                Else
                    ValDiameter = ValDiameter / 1000
                End If
                'читаем толщину стенки
                Dim WallT As Double = 0
                Dim strVallT As String = FuncPropertySet.FuncReadValueSetProperties(acFeatLine, "WALL_THICKNESS")
                If Val(strVallT) > 0 Then
                    WallT = Val(strVallT)
                Else
                    Dim optDb1 As PromptDoubleOptions = New PromptDoubleOptions("Введите толщину стенки трубы: ")
                    Dim resDb1 As PromptDoubleResult = ed.GetDouble(optDb1)
                    WallT = Math.Abs(resDb1.Value)
                End If
                WallT = WallT / 1000
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim acCirc1 As Circle = New Circle()
                Dim acCirc2 As Circle = New Circle()
                Dim acPline1 As Polyline = New Polyline()
                Dim acPline2 As Polyline = New Polyline()
                'читаем все данные
                Dim ArrayData As String(,) = Nothing
                Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(acFeatLine, ArrayData)
                Dim IDTablePS As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties(ArrayData(0, 0))
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                Dim PointColl As Point3dCollection = acFeatLine.GetPoints(ftype)
                If PointColl.Count > 1 Then
                    Using acPoly3d As Polyline3d = New Polyline3d()
                        acBlkTblRec.AppendEntity(acPoly3d)
                        tr.AddNewlyCreatedDBObject(acPoly3d, True)
                        For Each pt As Point3d In PointColl
                            If ValDiameter > 0 Then
                                Dim newPt As Point3d = New Point3d(pt.X, pt.Y, pt.Z - ValDiameter / 2)
                                Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(newPt)
                                    acPoly3d.AppendVertex(acPolVer3d)
                                    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                                End Using
                            ElseIf ValHeight > 0 Then
                                Dim newPt As Point3d = New Point3d(pt.X, pt.Y, pt.Z - ValHeight / 2)
                                Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(newPt)
                                    acPoly3d.AppendVertex(acPolVer3d)
                                    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                                End Using
                            End If
                        Next
                        'рисуем окружность 1 с внутренним диаметром
                        If ValDiameter > 0 Then
                            acBlkTblRec.AppendEntity(acCirc2)
                            tr.AddNewlyCreatedDBObject(acCirc2, True)
                            acCirc2.Diameter = ValDiameter
                            acCirc2.Center = acPoly3d.StartPoint
                            'рисуем вторую окружность с внешним диаметром
                            If WallT > 0 Then
                                acBlkTblRec.AppendEntity(acCirc1)
                                tr.AddNewlyCreatedDBObject(acCirc1, True)
                                acCirc1.Diameter = ValDiameter + WallT
                                acCirc1.Center = acPoly3d.StartPoint
                            End If
                        ElseIf ValHeight > 0 And ValWidth > 0 Then
                            'Dim angle As Double = acPoly3d.StartPoint.Convert2d(New Plane).GetVectorTo(acPoly3d.EndPoint.Convert2d(New Plane)).Angle + Math.PI / 2
                            'If angle >= Math.PI * 2 Then
                            '    angle -= Math.PI * 2
                            'End If
                            'Dim dx1 As Double = ValWidth * Math.Cos(angle)
                            'Dim dy1 As Double = ValWidth * Math.Sin(angle)
                            'Dim pt1 As Point3d = New Point3d(acPoly3d.StartPoint.X + dx1, acPoly3d.StartPoint.Y + dy1, acPoly3d.StartPoint.Z + ValHeight / 2)
                            'Dim pt3 As Point3d = New Point3d(acPoly3d.StartPoint.X + dx1, acPoly3d.StartPoint.Y + dy1, acPoly3d.StartPoint.Z - ValHeight / 2)
                            'Dim dx2 As Double = ValWidth * Math.Cos(angle + Math.PI)
                            'Dim dy2 As Double = ValWidth * Math.Sin(angle + Math.PI)
                            'Dim pt2 As Point3d = New Point3d(acPoly3d.StartPoint.X + dx2, acPoly3d.StartPoint.Y + dy2, acPoly3d.StartPoint.Z + ValHeight / 2)
                            'Dim pt4 As Point3d = New Point3d(acPoly3d.StartPoint.X + dx2, acPoly3d.StartPoint.Y + dy2, acPoly3d.StartPoint.Z - ValHeight / 2)
                            'Dim acPoly3d1 As Polyline3d = New Polyline3d()
                            'acBlkTblRec.AppendEntity(acPoly3d1)
                            'tr.AddNewlyCreatedDBObject(acPoly3d1, True)
                            'Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(pt1)
                            '    acPoly3d1.AppendVertex(acPolVer3d)
                            '    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                            'End Using
                            'Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(pt2)
                            '    acPoly3d1.AppendVertex(acPolVer3d)
                            '    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                            'End Using
                            'Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(pt4)
                            '    acPoly3d1.AppendVertex(acPolVer3d)
                            '    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                            'End Using
                            'Using acPolVer3d As PolylineVertex3d = New PolylineVertex3d(pt3)
                            '    acPoly3d1.AppendVertex(acPolVer3d)
                            '    tr.AddNewlyCreatedDBObject(acPolVer3d, True)
                            'End Using

                            'acPoly3d1.Closed = True

                            'Dim sob22 As SweepOptionsBuilder = New SweepOptionsBuilder()
                            'sob22.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
                            'sob22.BasePoint = acPoly3d.StartPoint
                            'sob22.Bank = True

                            'Dim sol11 As Solid3d = New Solid3d()
                            'sol11.CreateSweptSolid(acPline1, acPoly3d1, sob22.ToSweepOptions())
                            'acBlkTblRec.AppendEntity(sol11)
                            'tr.AddNewlyCreatedDBObject(sol11, True)
                            'GoTo line1
                            acBlkTblRec.AppendEntity(acPline2)
                            tr.AddNewlyCreatedDBObject(acPline2, True)
                            Dim pt1 As Point2d = New Point2d(acPoly3d.StartPoint.X - ValWidth / 2, acPoly3d.StartPoint.Y + ValHeight / 2)
                            Dim pt2 As Point2d = New Point2d(acPoly3d.StartPoint.X + ValWidth / 2, acPoly3d.StartPoint.Y + ValHeight / 2)
                            Dim pt3 As Point2d = New Point2d(acPoly3d.StartPoint.X + ValWidth / 2, acPoly3d.StartPoint.Y - ValHeight / 2)
                            Dim pt4 As Point2d = New Point2d(acPoly3d.StartPoint.X - ValWidth / 2, acPoly3d.StartPoint.Y - ValHeight / 2)
                            acPline2.AddVertexAt(0, pt1, 0, 0, 0)
                            acPline2.AddVertexAt(1, pt2, 0, 0, 0)
                            acPline2.AddVertexAt(2, pt3, 0, 0, 0)
                            acPline2.AddVertexAt(3, pt4, 0, 0, 0)
                            acPline2.Closed = True
                            Dim curUCSMatrix As Matrix3d = doc.Editor.CurrentUserCoordinateSystem
                            Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
                            Dim angle As Double = acPoly3d.StartPoint.Convert2d(New Plane).GetVectorTo(acPoly3d.EndPoint.Convert2d(New Plane)).Angle + Math.PI / 2
                            If angle >= Math.PI * 2 Then
                                angle -= Math.PI * 2
                            End If
                            acPline2.TransformBy(Matrix3d.Rotation(angle, curUCS.Zaxis, New Point3d(acPoly3d.StartPoint.X, acPoly3d.StartPoint.Y, 0)))
                            acPline2.Elevation = acPoly3d.StartPoint.Z
                            acBlkTblRec.AppendEntity(acPline1)
                            tr.AddNewlyCreatedDBObject(acPline1, True)
                            Dim pt5 As Point2d = New Point2d(acPoly3d.StartPoint.X - (ValWidth + WallT) / 2, acPoly3d.StartPoint.Y + (ValHeight + WallT) / 2)
                            Dim pt6 As Point2d = New Point2d(acPoly3d.StartPoint.X + (ValWidth + WallT) / 2, acPoly3d.StartPoint.Y + (ValHeight + WallT) / 2)
                            Dim pt7 As Point2d = New Point2d(acPoly3d.StartPoint.X + (ValWidth + WallT) / 2, acPoly3d.StartPoint.Y - (ValHeight + WallT) / 2)
                            Dim pt8 As Point2d = New Point2d(acPoly3d.StartPoint.X - (ValWidth + WallT) / 2, acPoly3d.StartPoint.Y - (ValHeight + WallT) / 2)
                            acPline1.AddVertexAt(0, pt5, 0, 0, 0)
                            acPline1.AddVertexAt(1, pt6, 0, 0, 0)
                            acPline1.AddVertexAt(2, pt7, 0, 0, 0)
                            acPline1.AddVertexAt(3, pt8, 0, 0, 0)
                            acPline1.Closed = True
                            acPline1.TransformBy(Matrix3d.Rotation(angle, curUCS.Zaxis, New Point3d(acPoly3d.StartPoint.X, acPoly3d.StartPoint.Y, 0)))
                            acPline1.Elevation = acPoly3d.StartPoint.Z
                        Else
                            MsgBox("YНет данных по диаметру трубы!!! Макро прервано")
                            Exit Sub
                        End If

                        Dim sob As SweepOptionsBuilder = New SweepOptionsBuilder()
                        sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
                        sob.BasePoint = acPoly3d.StartPoint
                        sob.Bank = True

                        'создаем 2 солида
                        Dim sol1 As Solid3d = New Solid3d()
                        Dim sol2 As Solid3d = New Solid3d()
                        If ValDiameter > 0 Then
                            sol1.CreateSweptSolid(acCirc1, acPoly3d, sob.ToSweepOptions())
                            acBlkTblRec.AppendEntity(sol1)
                            tr.AddNewlyCreatedDBObject(sol1, True)

                            sol2.CreateSweptSolid(acCirc2, acPoly3d, sob.ToSweepOptions())
                            acBlkTblRec.AppendEntity(sol2)
                            tr.AddNewlyCreatedDBObject(sol2, True)
                        Else

                            sol1.CreateSweptSolid(acPline1, acPoly3d, sob.ToSweepOptions())
                            acBlkTblRec.AppendEntity(sol1)
                            tr.AddNewlyCreatedDBObject(sol1, True)

                            sol2.CreateSweptSolid(acPline2, acPoly3d, sob.ToSweepOptions())
                            acBlkTblRec.AppendEntity(sol2)
                            tr.AddNewlyCreatedDBObject(sol2, True)
                        End If

                        'вычитаем солиды
                        If sol1.CheckInterference(sol2) = True And WallT > 0 Then
                            sol1.BooleanOperation(BooleanOperationType.BoolSubtract, sol2)
                        End If

                        'присваиваем расширенные данные
                        For i As Integer = 0 To ArrayData.GetUpperBound(1)
                            Dim NameField As String = ArrayData(1, i)
                            Dim val As String = ArrayData(2, i)
                            If IsNothing(val) = True Then
                                val = "-"
                            ElseIf val.Trim.Length = 0 Then
                                val = "-"
                            End If
                            If NameField Like "LEVEL1_DOWN" Then
                                If val Like "-" Then
                                    val = Math.Round(acFeatLine.StartPoint.Z - ValDiameter, 2)
                                End If
                            ElseIf NameField Like "LEVEL1_TOP" Then
                                If val Like "-" Then
                                    val = Math.Round(acFeatLine.StartPoint.Z, 2)
                                End If
                            ElseIf NameField Like "LEVEL2_DOWN" Then
                                If val Like "-" Then
                                    val = Math.Round(acFeatLine.EndPoint.Z - ValDiameter, 2)
                                End If
                            ElseIf NameField Like "LEVEL2_TOP" Then
                                If val Like "-" Then
                                    val = Math.Round(acFeatLine.EndPoint.Z, 2)
                                End If

                            End If
                            ArrayData(2, i) = val
                        Next i
                        If IsArray(ArrayData) = True Then
                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(sol1, IDTablePS)
                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(sol1, ArrayData)
                        End If
                        sol1.Layer = nameLayer
                        acCirc1.Erase()
                        acCirc2.Erase()
                        acPoly3d.Erase()
                    End Using
                End If
line1:
                tr.Commit()
            End If
        End Using
    End Sub



    '=============================================================================================
    '1. Первое действие, вставка объектов сивил 3D (вставка колодца и трубы)
    '=============================================================================================
    <CommandMethod("ADSK-CreatePipeNeywork")>
    Public Sub CreatyPipeNeywork1()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim formDrawPipeLine As FormDrawPipeNetwork = New FormDrawPipeNetwork
        formDrawPipeLine.ShowDialog()
        If formDrawPipeLine.formShow = False Then
            Exit Sub
        End If
        'величина удлиннения трубы
        Dim dLPipe As Double = formDrawPipeLine.NumericUpDown3.Value
        'вертикальное смещение (от дна колодца до низа трубы)
        Dim dhPipe As Double = formDrawPipeLine.NumericUpDown2.Value
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim oNetworkId As ObjectId = Nothing
            Dim oNetwork As Network = Nothing
            Dim NameNetwork As String = formDrawPipeLine.ComboBox1.Text
            Dim oPipeNetworkIds As ObjectIdCollection = CivilDoc.GetPipeNetworkIds()
            If (oPipeNetworkIds Is Nothing) Then
                Exit Sub
            End If
            If oPipeNetworkIds.Count > 0 Then
                For Each IDNetwork As ObjectId In oPipeNetworkIds
                    If IDNetwork.IsNull = False Then
                        Dim TempNetwork As Network = tr.GetObject(IDNetwork, OpenMode.ForWrite)
                        If TempNetwork.Name Like NameNetwork Then
                            oNetworkId = TempNetwork.Id
                        End If
                    End If
                Next
            End If
            If IsNothing(oNetworkId) = True Then
                oNetworkId = Network.Create(CivilDoc, formDrawPipeLine.ComboBox1.Text)
            End If
            If oNetworkId.IsNull = True Then
                oNetworkId = Network.Create(CivilDoc, formDrawPipeLine.ComboBox1.Text)
            End If
            Try
                oNetwork = tr.GetObject(oNetworkId, OpenMode.ForWrite)
            Catch
                Exit Sub
            End Try
            '=======================================================================================
            'стиль
            'каталог трубопроводной сети
            Dim NamePartsList As String = formDrawPipeLine.ComboBox2.Text
            Dim oPartsListId As ObjectId = CivilDoc.Styles.PartsListSet.Item(NamePartsList) 'Standard PartsList
            Dim oPartsList As PartsList = tr.GetObject(oPartsListId, OpenMode.ForWrite)

            Dim NameoldPipe As String = formDrawPipeLine.ComboBox3.Text
            Dim IDoidPipe As ObjectId = oPartsList.Item(NameoldPipe)
            Dim opfPipe As PartFamily = tr.GetObject(IDoidPipe, OpenMode.ForWrite)

            Dim NameSizePipe As String = formDrawPipeLine.ComboBox4.Text
            Dim idPSizePipe As ObjectId = opfPipe.Item(NameSizePipe)
            Dim PSizePipe As PartSize = tr.GetObject(idPSizePipe, OpenMode.ForWrite)

            Dim NameopfStructure As String = formDrawPipeLine.ComboBox5.Text
            Dim IDoidStruct As ObjectId = oPartsList.Item(NameopfStructure)
            Dim opfStructure As PartFamily = tr.GetObject(IDoidStruct, OpenMode.ForWrite)

            Dim NameSizeStruct As String = formDrawPipeLine.ComboBox6.Text
            Dim idPSizeStruct As ObjectId = opfStructure.Item(NameSizeStruct)
            Dim PSizeStruct As PartSize = tr.GetObject(idPSizeStruct, OpenMode.ForWrite)

            Dim AlignProject As Alignment = Nothing
            Dim AlignCool As ObjectIdCollection = CivilDoc.GetAlignmentIds
            If AlignCool.Count > 0 Then
                For Each idAlign As ObjectId In AlignCool
                    Dim Align As Autodesk.Civil.DatabaseServices.Alignment = idAlign.GetObject(OpenMode.ForRead)
                    If Align.Name Like formDrawPipeLine.ComboBox7.Text Then
                        AlignProject = Align
                        Exit For
                    End If
                Next
            End If
            Dim SurfProject As TinSurface = Nothing
            Dim SurfCool As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If SurfCool.Count > 0 Then
                For Each idSurf As ObjectId In SurfCool
                    Dim Surf As Autodesk.Civil.DatabaseServices.TinSurface = idSurf.GetObject(OpenMode.ForRead)
                    If Surf.Name Like formDrawPipeLine.ComboBox8.Text Then
                        SurfProject = Surf
                        Exit For
                    End If
                Next
            End If

            Dim SurfProject2 As TinSurface = Nothing
            Dim SurfCool2 As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If SurfCool2.Count > 0 Then
                For Each idSurf2 As ObjectId In SurfCool2
                    Dim Surf2 As Autodesk.Civil.DatabaseServices.TinSurface = idSurf2.GetObject(OpenMode.ForRead)
                    If Surf2.Name Like formDrawPipeLine.ComboBox9.Text Then
                        SurfProject2 = Surf2
                        Exit For
                    End If
                Next
            End If
            If IsNothing(opfPipe) = False Then
                If opfPipe.Id.IsNull = False Then
                    'создание сети из выбранной структурной линии
                    If formDrawPipeLine.RadioButton1.Checked = True Then
                        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                        If acSSPrompt.Status = PromptStatus.OK Then
                            Dim acSSet As SelectionSet = acSSPrompt.Value
                            For Each acSSObj As SelectedObject In acSSet
                                If Not IsDBNull(acSSObj) Then
                                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                    If Not IsDBNull(acEnt) Then
                                        If TypeOf acEnt Is FeatureLine Then
                                            Dim acFLine As FeatureLine = acEnt
                                            Dim colPt As Point3dCollection = acFLine.GetPoints(FeatureLinePointType.PIPoint)
                                            If colPt.Count > 1 Then
                                                For i As Integer = 0 To colPt.Count - 2
                                                    Dim line As LineSegment3d = New LineSegment3d(colPt.Item(i), colPt.Item(i + 1))
                                                    Dim oidNewPipe As ObjectId = Nothing
                                                    oNetwork.AddLinePipe(opfPipe.Id, PSizePipe.Id, line, oidNewPipe, False)
                                                    Dim oidNewStructure As ObjectId = ObjectId.Null
                                                    oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, colPt.Item(i), 0, oidNewStructure, False)
                                                    oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, colPt.Item(i + 1), 0, oidNewStructure, False)
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        'из файла
                    ElseIf formDrawPipeLine.RadioButton2.Checked = True Then
                        'Выбираем файл Excel
                        Dim opFile As New System.Windows.Forms.OpenFileDialog()
                        opFile.InitialDirectory = "d:\ЗАХАРОВ\НГ\Водоотвод_Тест\"
                        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
                        opFile.RestoreDirectory = True
                        opFile.ShowDialog()
                        Dim fFile1 As String = opFile.FileName
                        If IsNothing(fFile1) = True Then Exit Sub
                        If fFile1.Trim.Length = 0 Then Exit Sub
                        Dim xlApp As Excel.Application = Nothing
                        Dim xlBook As Excel.Workbook = Nothing
                        Dim xlSheets As Excel.Worksheet = Nothing
                        xlApp = CreateObject("Excel.Application")
                        xlBook = xlApp.Workbooks.Open(fFile1)
                        Dim xlsheet As Excel.Worksheet
                        xlsheet = xlBook.Worksheets.Item(formDrawPipeLine.NumericUpDown5.Value) 'делаем ссылку на этот лист
                        Dim CountErr As Integer = 100
                        Dim ArrayDataExcel As Double(,) = Nothing
                        Dim ArrayDataExcel2 As String(,) = Nothing
                        Dim CountArrayDataExcel As Integer = 0
                        If formDrawPipeLine.CheckBox1.Checked = False Then
                            For indRows As Integer = 6 To xlsheet.Rows.Count - 1
                                If CountErr = 0 Then Exit For
                                Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
                                If IsNumeric(NumbStr) = False Then
                                    CountErr -= 1
                                    Continue For
                                End If
                                Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
                                Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
                                Dim LPipe As String = xlsheet.Cells(indRows, 4).value + dLPipe 'длина трубы
                                Dim iPipe As String = xlsheet.Cells(indRows, 5).value 'уклон
                                Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
                                Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2
                                Dim H3 As String = xlsheet.Cells(indRows, 8).value 'H3
                                Dim H4 As String = xlsheet.Cells(indRows, 9).value 'H4
                                Dim H5 As String = xlsheet.Cells(indRows, 10).value 'H5
                                Dim TypeG As String = xlsheet.Cells(indRows, 11).value 'тип гасителя
                                If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
                                    CountErr -= 1
                                    Continue For
                                ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
                                    CountErr -= 1
                                    Continue For
                                Else
                                    ReDim Preserve ArrayDataExcel(10, CountArrayDataExcel)
                                    If IsNothing(StartLeftPKStr) = False Then
                                        If StartLeftPKStr.Trim.Length > 0 Then
                                            ArrayDataExcel(0, CountArrayDataExcel) = Val(StartLeftPKStr)
                                            ArrayDataExcel(1, CountArrayDataExcel) = -777
                                        End If
                                    End If
                                    If IsNothing(StartRightPKStr) = False Then
                                        If StartRightPKStr.Trim.Length > 0 Then
                                            ArrayDataExcel(1, CountArrayDataExcel) = Val(StartRightPKStr)
                                            ArrayDataExcel(0, CountArrayDataExcel) = -777
                                        End If
                                    End If
                                    ArrayDataExcel(2, CountArrayDataExcel) = Val(LPipe)
                                    ArrayDataExcel(3, CountArrayDataExcel) = Val(iPipe)
                                    ArrayDataExcel(4, CountArrayDataExcel) = Val(H1)
                                    ArrayDataExcel(5, CountArrayDataExcel) = Val(H2)
                                    ArrayDataExcel(6, CountArrayDataExcel) = Val(H3)
                                    ArrayDataExcel(7, CountArrayDataExcel) = Val(H4)
                                    ArrayDataExcel(8, CountArrayDataExcel) = Val(H5)
                                    ArrayDataExcel(9, CountArrayDataExcel) = Val(TypeG)
                                    CountArrayDataExcel += 1
                                    CountErr = 100
                                End If
                            Next
                        Else
                            For indRows As Integer = 3 To xlsheet.Rows.Count - 1
                                If CountErr = 0 Then Exit For
                                Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
                                If IsNothing(NumbStr) = True Then
                                    CountErr -= 1
                                    Continue For
                                End If

                                Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
                                Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
                                Dim iPipe As String = xlsheet.Cells(indRows, 4).value 'уклон
                                Dim offStructure As String = xlsheet.Cells(indRows, 5).value 'смещение колодца
                                Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
                                Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2

                                If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
                                    CountErr -= 1
                                    Continue For
                                ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
                                    CountErr -= 1
                                    Continue For
                                Else
                                    ReDim Preserve ArrayDataExcel2(10, CountArrayDataExcel)
                                    If IsNothing(StartLeftPKStr) = False Then
                                        If StartLeftPKStr.Trim.Length > 0 Then
                                            ArrayDataExcel2(0, CountArrayDataExcel) = StartLeftPKStr
                                            ArrayDataExcel2(1, CountArrayDataExcel) = "-"
                                        End If
                                    End If
                                    If IsNothing(StartRightPKStr) = False Then
                                        If StartRightPKStr.Trim.Length > 0 Then
                                            ArrayDataExcel2(1, CountArrayDataExcel) = Val(StartRightPKStr)
                                            ArrayDataExcel2(0, CountArrayDataExcel) = "-"
                                        End If
                                    End If
                                    ArrayDataExcel2(2, CountArrayDataExcel) = offStructure
                                    ArrayDataExcel2(3, CountArrayDataExcel) = iPipe
                                    ArrayDataExcel2(4, CountArrayDataExcel) = H1
                                    ArrayDataExcel2(5, CountArrayDataExcel) = H2
                                    ArrayDataExcel2(6, CountArrayDataExcel) = ""
                                    ArrayDataExcel2(7, CountArrayDataExcel) = ""
                                    ArrayDataExcel2(8, CountArrayDataExcel) = ""
                                    ArrayDataExcel2(9, CountArrayDataExcel) = NumbStr
                                    CountArrayDataExcel += 1
                                    CountErr = 100
                                End If
                            Next
                        End If
                        xlBook.Close()
                        xlApp.Quit()
                        If formDrawPipeLine.CheckBox1.Checked = False Then
                            If IsArray(ArrayDataExcel) = True Then
                                For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
                                    Dim LPipe As Double = ArrayDataExcel(2, i)
                                    Dim off1 As Double = formDrawPipeLine.NumericUpDown2.Value
                                    Dim off2 As Double = formDrawPipeLine.NumericUpDown2.Value + LPipe
                                    Dim PK As Double = 0
                                    If ArrayDataExcel(0, i) <> -777 Then
                                        PK = ArrayDataExcel(0, i)
                                        off1 = off1 * -1
                                        off2 = off2 * -1
                                    Else
                                        PK = ArrayDataExcel(1, i)
                                    End If
                                    'отметки входа и выхода трубы
                                    Dim H2 As Double = ArrayDataExcel(5, i)
                                    Dim H3 As Double = ArrayDataExcel(6, i)
                                    H2 = H2 + formDrawPipeLine.NumericUpDown1.Value
                                    H3 = H3 + formDrawPipeLine.NumericUpDown1.Value
                                    'отметка верха колодца
                                    Dim H1 As Double = ArrayDataExcel(4, i)
                                    'глубина колодца
                                    Dim dhStruct As Double = Math.Abs(H1 - ArrayDataExcel(8, i))

                                    Dim XStartPipe As Double = -1
                                    Dim YStartPipe As Double = -1
                                    Dim XEndPipe As Double = -1
                                    Dim YEndPipe As Double = -1
                                    Try
                                        AlignProject.PointLocation(PK, off1, XStartPipe, YStartPipe)
                                        AlignProject.PointLocation(PK, off2, XEndPipe, YEndPipe)
                                    Catch ex As System.ArgumentOutOfRangeException
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    End Try
                                    If XStartPipe <> -1 And YStartPipe <> -1 And XEndPipe <> -1 And YEndPipe <> -1 Then
                                        Try
                                            Dim ptStartPipe As Point3d = New Point3d(XStartPipe, YStartPipe, H2)
                                            Dim ptEndPipe As Point3d = New Point3d(XEndPipe, YEndPipe, H3)
                                            Dim angleStruct As Double = ptStartPipe.Convert2d(New Plane).GetVectorTo(ptEndPipe.Convert2d(New Plane)).Angle
                                            Dim line As LineSegment3d = New LineSegment3d(ptStartPipe, ptEndPipe)
                                            Dim oidNewPipe As ObjectId = Nothing
                                            oNetwork.AddLinePipe(opfPipe.Id, PSizePipe.Id, line, oidNewPipe, False)
                                            If oidNewPipe.IsNull = False Then
                                                Dim acPipe As Pipe = tr.GetObject(oidNewPipe, OpenMode.ForWrite)
                                                acPipe.RefAlignmentId = AlignProject.Id
                                                acPipe.HoldOnResizeType = HoldOnResizeType.Invert
                                                acPipe.StartPoint = ptStartPipe
                                                acPipe.EndPoint = ptEndPipe
                                                Dim oData As PartDataField() = acPipe.PartData.GetAllDataFields()
                                                For Each idp As PartDataField In oData
                                                    Try
                                                        Dim Name As String = idp.Name
                                                        Dim Desk As String = idp.Description
                                                        Dim NameFieldDimPipe As String = ""
                                                        If Desk Like NameFieldDimPipe Then
                                                            If idp.IsFromList Then
                                                                Dim Record As PartDataRecord = acPipe.PartData
                                                                idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                                Dim valL As PartDataList = idp.ValueList
                                                                If valL.Count > 0 Then
                                                                    For i1 As Integer = 0 To valL.Count - 1
                                                                        Dim obj As Object = valL.Item(i1)
                                                                        Dim a As Integer = 0
                                                                    Next
                                                                End If

                                                            Else
                                                                Dim Record As PartDataRecord = acPipe.PartData
                                                                idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                                If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
                                                                    idp.Value = formDrawPipeLine.NumericUpDown1.Value
                                                                Else
                                                                    idp.Value = formDrawPipeLine.NumericUpDown1.Value
                                                                End If
                                                                acPipe.PartData = Record
                                                            End If
                                                        End If
                                                    Catch ex As InvalidOperationException
                                                    Catch ex As ArgumentException
                                                    End Try
                                                Next

                                            End If
                                            Dim ptStartStruct As Point3d = New Point3d(XStartPipe, YStartPipe, H1)
                                            Dim oidNewStructure As ObjectId = ObjectId.Null
                                            oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, ptStartStruct, 0, oidNewStructure, False)
                                            If oidNewStructure.IsNull = False Then
                                                Dim acStruct As [Structure] = tr.GetObject(oidNewStructure, OpenMode.ForWrite)
                                                acStruct.SumpDepth = dhStruct
                                                Dim oData As PartDataField() = acStruct.PartData.GetAllDataFields()
                                                acStruct.Rotation = angleStruct
                                                acStruct.RefAlignmentId = AlignProject.Id
                                                For Each idp As PartDataField In oData
                                                    Try
                                                        Dim Name As String = idp.Name
                                                        Dim Desk As String = idp.Description
                                                        Dim NameFieldDimPipe As String = "" 'FormDtawPipeLine.TextBox4.Text.Trim
                                                        If Desk Like NameFieldDimPipe Then
                                                            If idp.IsFromList Then
                                                                Dim Record As PartDataRecord = acStruct.PartData
                                                                idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                                Dim valL As PartDataList = idp.ValueList
                                                                If valL.Count > 0 Then
                                                                    For i1 As Integer = 0 To valL.Count - 1
                                                                        Dim obj As Object = valL.Item(i1)
                                                                        Dim a As Integer = 0
                                                                    Next
                                                                End If

                                                            Else
                                                                Dim Record As PartDataRecord = acStruct.PartData
                                                                idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                                If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
                                                                    idp.Value = dhStruct
                                                                Else
                                                                    idp.Value = dhStruct
                                                                End If
                                                                acStruct.PartData = Record
                                                            End If
                                                        End If
                                                    Catch ex As InvalidOperationException
                                                    Catch ex As ArgumentException
                                                    End Try
                                                Next
                                            End If
                                        Catch ex As System.Exception
                                        End Try
                                    End If
                                Next i
                            End If
                        Else
                            If IsArray(ArrayDataExcel2) = True Then
                                For i As Integer = 0 To ArrayDataExcel2.GetUpperBound(1)
                                    Dim PKStr As String = ArrayDataExcel2(0, i)
                                    Dim k As Integer = -1
                                    If PKStr Like "-" Then
                                        PKStr = ArrayDataExcel2(1, i)
                                        k = 1
                                    End If
                                    PKStr.Replace(",", ".")
                                    Dim PKStructure As Double = Val(PKStr)

                                    Dim offStr As String = ArrayDataExcel2(2, i)
                                    offStr = offStr.Replace(",", ".")
                                    Dim offStructure As Double = Val(offStr)
                                    Dim XCentreStructure As Double = 0
                                    Dim YCentreStructure As Double = 0
                                    Try
                                        AlignProject.PointLocation(PKStructure, k * offStructure, XCentreStructure, YCentreStructure)
                                    Catch ex As System.ArgumentOutOfRangeException
                                        PKStructure += 0.01

                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        PKStructure += 0.01
                                    End Try
                                    Dim H1 As Double = -1
                                    Try
                                        H1 = SurfProject.FindElevationAtXY(XCentreStructure, YCentreStructure)
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    End Try
                                    If H1 = -1 Then
                                        Try
                                            H1 = SurfProject2.FindElevationAtXY(XCentreStructure, YCentreStructure)
                                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                            Continue For
                                        End Try
                                    End If
                                    Dim insPointStructure As Point3d = New Point3d(XCentreStructure, YCentreStructure, H1)

                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'вычисляем начало трубы
                                    Dim strHPipe As String = ArrayDataExcel2(5, i)
                                    strHPipe = strHPipe.Replace(",", ".")
                                    Dim hhPipe As Double = Val(strHPipe)
                                    If formDrawPipeLine.CheckBox2.Checked = True Then
                                        hhPipe = H1 - formDrawPipeLine.NumericUpDown4.Value / 1000
                                    End If
                                    Dim hStructure As Double = (H1 - hhPipe) + formDrawPipeLine.NumericUpDown2.Value / 1000
                                    Dim startPointPipe As Point3d = New Point3d(XCentreStructure, YCentreStructure, hhPipe)
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'временно удлинняем линию
                                    Dim XEndPipe As Double = 0
                                    Dim YEndPipe As Double = 0
                                    Try
                                        If offStructure < 0 Then
                                            k = -1
                                        End If
                                        AlignProject.PointLocation(PKStructure, k * (offStructure + 50), XEndPipe, YEndPipe)
                                    Catch ex As System.ArgumentOutOfRangeException
                                        Continue For
                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                        Continue For
                                    End Try
                                    Dim iPipe As Double = Val(ArrayDataExcel2(3, i)) / 1000
                                    Dim dh As Double = (offStructure + 50) * iPipe
                                    Dim endPointPipe As Point3d = New Point3d(XEndPipe, YEndPipe, hhPipe - dh)
                                    'обрезаем линию по поверхности
                                    Dim V3d As Vector3d = startPointPipe.GetVectorTo(endPointPipe)
                                    Dim PtIntersect As Point3d = New Point3d(-1, -1, -1)
                                    Try
                                        PtIntersect = SurfProject.GetIntersectionPoint(startPointPipe, V3d)
                                    Catch ex As System.ArgumentException
                                    End Try
                                    If PtIntersect.X = -1 And PtIntersect.Y = -1 And PtIntersect.Z = -1 Then
                                        Try
                                            PtIntersect = SurfProject2.GetIntersectionPoint(startPointPipe, V3d)
                                            endPointPipe = PtIntersect
                                        Catch ex As System.ArgumentException
                                        End Try
                                    End If
                                    'удлинням линию
                                    Dim LenPipe As Double = formDrawPipeLine.NumericUpDown3.Value / 1000
                                    Dim acLine As Line = New Line(startPointPipe, endPointPipe)
                                    Dim boolExt As Boolean = MathFunction.FuncExtendLine(acLine, 0, LenPipe, 3)
                                    endPointPipe = acLine.EndPoint
                                    Try
                                        Dim ptStartPipe As Point3d = startPointPipe
                                        Dim ptEndPipe As Point3d = endPointPipe
                                        Dim angleStruct As Double = ptStartPipe.Convert2d(New Plane).GetVectorTo(ptEndPipe.Convert2d(New Plane)).Angle
                                        Dim line As LineSegment3d = New LineSegment3d(ptStartPipe, ptEndPipe)
                                        Dim oidNewPipe As ObjectId = Nothing
                                        oNetwork.AddLinePipe(opfPipe.Id, PSizePipe.Id, line, oidNewPipe, False)
                                        If oidNewPipe.IsNull = False Then
                                            Dim acPipe As Pipe = tr.GetObject(oidNewPipe, OpenMode.ForWrite)
                                            acPipe.RefAlignmentId = AlignProject.Id
                                            acPipe.HoldOnResizeType = HoldOnResizeType.Invert
                                            Dim d As Double = formDrawPipeLine.NumericUpDown1.Value / 1000
                                            acPipe.StartPoint = New Point3d(ptStartPipe.X, ptStartPipe.Y, ptStartPipe.Z + d / 2)
                                            acPipe.EndPoint = New Point3d(ptEndPipe.X, ptEndPipe.Y, ptEndPipe.Z + d / 2)
                                            Dim oData As PartDataField() = acPipe.PartData.GetAllDataFields()
                                            For Each idp As PartDataField In oData
                                                Try
                                                    Dim Name As String = idp.Name
                                                    Dim Desk As String = idp.Description
                                                    Dim NameFieldDimPipe As String = ""
                                                    If Desk Like NameFieldDimPipe Then
                                                        If idp.IsFromList Then
                                                            Dim Record As PartDataRecord = acPipe.PartData
                                                            idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                            Dim valL As PartDataList = idp.ValueList
                                                            If valL.Count > 0 Then
                                                                For i1 As Integer = 0 To valL.Count - 1
                                                                    Dim obj As Object = valL.Item(i1)
                                                                    Dim a As Integer = 0
                                                                Next
                                                            End If
                                                        Else
                                                            Dim Record As PartDataRecord = acPipe.PartData
                                                            idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                            If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
                                                                idp.Value = formDrawPipeLine.NumericUpDown1.Value / 1000
                                                            Else
                                                                idp.Value = formDrawPipeLine.NumericUpDown1.Value / 1000
                                                            End If
                                                            acPipe.PartData = Record
                                                        End If
                                                    End If
                                                Catch ex As InvalidOperationException
                                                Catch ex As ArgumentException
                                                End Try
                                            Next

                                        End If
                                        Dim ptStartStruct As Point3d = insPointStructure
                                        Dim oidNewStructure As ObjectId = ObjectId.Null
                                        oNetwork.AddStructure(opfStructure.Id, PSizeStruct.Id, ptStartStruct, 0, oidNewStructure, False)
                                        If oidNewStructure.IsNull = False Then
                                            Dim acStruct As [Structure] = tr.GetObject(oidNewStructure, OpenMode.ForWrite)
                                            acStruct.Name = ArrayDataExcel2(9, i)
                                            acStruct.SumpDepth = formDrawPipeLine.NumericUpDown4.Value / 1000
                                            Dim oData As PartDataField() = acStruct.PartData.GetAllDataFields()
                                            acStruct.Rotation = angleStruct
                                            acStruct.RefAlignmentId = AlignProject.Id
                                            For Each idp As PartDataField In oData
                                                Try
                                                    Dim Name As String = idp.Name
                                                    Dim Desk As String = idp.Description
                                                    Dim NameFieldDimPipe As String = "" 'FormDtawPipeLine.TextBox4.Text.Trim
                                                    If Desk Like NameFieldDimPipe Then
                                                        If idp.IsFromList Then
                                                            Dim Record As PartDataRecord = acStruct.PartData
                                                            idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                            Dim valL As PartDataList = idp.ValueList
                                                            If valL.Count > 0 Then
                                                                For i1 As Integer = 0 To valL.Count - 1
                                                                    Dim obj As Object = valL.Item(i1)
                                                                    Dim a As Integer = 0
                                                                Next
                                                            End If

                                                        Else
                                                            Dim Record As PartDataRecord = acStruct.PartData
                                                            idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                            If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
                                                                idp.Value = acStruct.SumpDepth
                                                            Else
                                                                idp.Value = acStruct.SumpDepth
                                                            End If
                                                            acStruct.PartData = Record
                                                        End If
                                                    End If
                                                Catch ex As InvalidOperationException
                                                Catch ex As ArgumentException
                                                End Try
                                            Next
                                        End If
                                    Catch ex As System.Exception
                                    End Try
                                Next i
                            End If
                        End If
                    End If
                End If
            End If
            tr.Commit()
        End Using
    End Sub
    ''=============================================================================================
    ''2. Замена колодца Civil3d на пользовательские колодцы
    ''=============================================================================================
    <CommandMethod("ADSK-InsertStructuresBLK")>
    Public Sub InsertStructuresBLK()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim formInsertBlk As FormInsertBlkToStructure = New FormInsertBlkToStructure
        formInsertBlk.ShowDialog()
        If formInsertBlk.boolShow = False Then
            Exit Sub
        End If
        Dim NameObject As String = formInsertBlk.ComboBox1.Text
        If IsNothing(NameObject) = True Then
            Exit Sub
        End If
        'Dim PutchFile As String = "k:\Civil3DToolsUtility\FileResources\3d_Object\Structures\Колодцы смотровые подземных коммуникаций.dwg"
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            If acBlkTbl.Has(NameObject) = True Then
                Dim IDBlk As ObjectId = acBlkTbl.Item(NameObject)
            Else
                MsgBox("Блок не найден. Макро превано.")
                Exit Sub
            End If
            Dim nameLayer As String = "0"
            If formInsertBlk.CheckBox2.Checked = False Then
                Dim tempNameLayer As String = formInsertBlk.TextBox1.Text
                If tempNameLayer.Trim.Length > 0 Then
                    Dim idLayer As ObjectId = AcadFunction.FuncAddLayer(tempNameLayer)
                    If IsNothing(idLayer) = False Then
                        If idLayer.IsNull = False Then
                            nameLayer = tempNameLayer
                        End If
                    End If
                End If
            End If
            'копируем определение блока
            'If IsNothing(NameObject) = False Then
            '    If NameObject.Trim.Length > 0 Then
            '        If acBlkTbl.Has(NameObject) = False Then
            '            Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFile, NameObject)
            '        End If
            '    End If
            'End If

            'находим колодцы
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is [Structure] Then
                                Dim Struct As [Structure] = acEnt
                                Dim structAngle As Double = Struct.Rotation
                                If formInsertBlk.CheckBox1.Checked = True Then
                                    structAngle = Struct.Rotation + Math.PI
                                End If
                                Dim PosPoint As Point3d = New Point3d(Struct.Position.X, Struct.Position.Y, Struct.Position.Z + formInsertBlk.NumericUpDown1.Value)
                                Dim pk As Double = Struct.Station
                                Dim PKStr As String = FuncFormatZN.FuncFormatPK(pk, 2)
                                'Struct.Erase()
                                Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(PosPoint, NameObject, nameLayer, 7, 1, 20, structAngle, 0)
                                If formInsertBlk.CheckBox2.Checked = True Then
                                    If IsNothing(acBlkObj) = False Then
                                        If acBlkObj.Id.IsNull = False Then
                                            Dim ObjColl As DBObjectCollection = New DBObjectCollection
                                            acBlkObj.Explode(ObjColl)
                                            If ObjColl.Count > 0 Then
                                                For Each db As Autodesk.AutoCAD.DatabaseServices.DBObject In ObjColl
                                                    Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = db
                                                    acBlkTblRec.AppendEntity(Ent)
                                                    acTrans.AddNewlyCreatedDBObject(Ent, True)
                                                    Dim boolnameTbl As Boolean = FuncPropertySet.FuncFindPSetTable(db, "PROJECT_PIPENETWORK")
                                                    If boolnameTbl = True Then
                                                        Dim ArrayODRec As String(,) = Nothing
                                                        ReDim Preserve ArrayODRec(3, 0)
                                                        ArrayODRec(0, 0) = "PROJECT_PIPENETWORK"
                                                        ArrayODRec(1, 0) = "STATION"
                                                        ArrayODRec(2, 0) = PKStr
                                                        Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(db, ArrayODRec)
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub
    ''=============================================================================================
    ''3. Замена труб Civil3d на солид, выдавливание трубы и построекние портальной стенки
    ''=============================================================================================
    <CommandMethod("ADSK-ReplacePipe")>
    Public Sub ReplacePipe()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'находим трубы
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Pipe Then
                                Dim acPipe As Pipe = acEnt
                                Dim PosPoint1 As Point3d = acPipe.StartPoint
                                Dim PosPoint2 As Point3d = acPipe.EndPoint
                                Dim acSolid3D As Solid3d = acPipe.Solid3dBody
                                acBlkTblRec.AppendEntity(acSolid3D)
                                acTrans.AddNewlyCreatedDBObject(acSolid3D, True)
                                'acPipe.Erase()
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub
    ''=============================================================================================
    ''4. Вставка портальной стенки
    ''=============================================================================================
    <CommandMethod("ADSK-InsertPortalVolume")>
    Public Sub InsertPortalVolume()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim formWall As FormInsertPortalWall = New FormInsertPortalWall
            formWall.ShowDialog()
            If formWall.boolShow = False Then
                Exit Sub
            End If
            Dim NameObject As String = formWall.ComboBox1.Text
            If IsNothing(NameObject) = True Then
                Exit Sub
            End If
            Dim SurfProject As TinSurface = Nothing
            Dim SurfCool As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If SurfCool.Count > 0 Then
                For Each idSurf As ObjectId In SurfCool
                    Dim Surf As Autodesk.Civil.DatabaseServices.TinSurface = idSurf.GetObject(OpenMode.ForRead)
                    If Surf.Name Like formWall.ComboBox2.Text Then
                        SurfProject = Surf
                        Exit For
                    End If
                Next
            End If

            If acBlkTbl.Has(NameObject) = True Then
                Dim IDBlk As ObjectId = acBlkTbl.Item(NameObject)
            Else
                MsgBox("Блок не найден. Макро превано.")
                Exit Sub
            End If
            Dim nameLayer As String = "0"
            If formWall.CheckBox1.Checked = False Then
                Dim tempNameLayer As String = formWall.TextBox1.Text
                If tempNameLayer.Trim.Length > 0 Then
                    Dim idLayer As ObjectId = AcadFunction.FuncAddLayer(tempNameLayer)
                    If IsNothing(idLayer) = False Then
                        If idLayer.IsNull = False Then
                            nameLayer = tempNameLayer
                        End If
                    End If
                End If
            End If
            Dim rPipe As Double = formWall.NumericUpDown1.Value / 2000
            'находим трубы
            Dim acSSPrompt1 As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt1.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt1.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Pipe Then
                                Dim acPipe As Pipe = acEnt
                                Dim PosPoint1 As Point3d = New Point3d(acPipe.StartPoint.X, acPipe.StartPoint.Y, acPipe.StartPoint.Z + rPipe)
                                Dim PosPoint2 As Point3d = New Point3d(acPipe.EndPoint.X, acPipe.EndPoint.Y, acPipe.EndPoint.Z + rPipe)
                                Dim acLine As Line = New Line(PosPoint2, PosPoint1)
                                Dim vPt As Vector3d = PosPoint1.GetVectorTo(PosPoint2)
                                Dim PtIntersect As Point3d = New Point3d(-1, -1, -1)
                                Try
                                    PtIntersect = SurfProject.GetIntersectionPoint(PosPoint1, vPt)
                                Catch ex As System.ArgumentException
                                End Try
                                If PtIntersect.X <> -1 And PtIntersect.Y <> -1 And PtIntersect.Z <> -1 Then
                                    Dim insPoint As Point3d = New Point3d(PtIntersect.X, PtIntersect.Y, PtIntersect.Z + 0.35)
                                    Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(insPoint, NameObject, nameLayer, 7, 1, 20, acLine.Angle + Math.PI / 2, 0)
                                End If

                                '    For i As Double = 0 To acLine.Length - 0.01 Step 0.01
                                '        Dim TempPt As Point3d = acLine.GetPointAtDist(i)
                                '        Try
                                '            Dim HTin As Double = TinSurface.FindElevationAtXY(TempPt.X, TempPt.Y)
                                '            Dim dh As Double = Math.Round(HTin - TempPt.Z, 2)
                                '            If dh >= 0.42 Then
                                '                Dim TempPt1 As Point3d = acLine.GetPointAtDist(i - 0.15)
                                '                Dim TempPt2 As Point3d = New Point3d(TempPt1.X, TempPt1.Y, TempPt1.Z + 0.42)
                                '                Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(TempPt2, NameObject, "0", 7, 1, 20, acLine.Angle + Math.PI / 2, 0)
                                '                Exit For
                                '            End If
                                '        Catch ex As Autodesk.Civil.PointNotOnEntityException
                                '        End Try

                                '    Next i
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub
    ''=============================================================================================
    ''5. Вставки песчаной подушкм
    ''=============================================================================================
    <CommandMethod("ADSK-InsertPesPodg")>
    Public Sub InsertPesPodg()
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
        Dim Ed As Editor = docCol.MdiActiveDocument.Editor
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ArrayID As ObjectId(,) = Nothing
        Dim CountArrayID As Integer = 0
        Dim MaxHeight As Double = 0.9
        Dim MinHeight As Double = 0.4
        Dim TopHeight As Double = 0.67
        Dim DownHeight As Double = -0.32
        Dim OffsetOtkos As Double = 0.1
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'выбираем поверхность
            Dim TinSurface As TinSurface = Nothing
            Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите поверхность: ")
            If Not (acSSPrompt.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            TinSurface = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead)
            'находим трубы
            Dim acSSPrompt1 As PromptSelectionResult = acDoc.Editor.GetSelection()
            If acSSPrompt1.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt1.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeOf acEnt Is Pipe Then
                                Dim acPipe As Pipe = acEnt
                                Dim PosPoint1 As Point3d = acPipe.StartPoint
                                Dim PosPoint2 As Point3d = acPipe.EndPoint
                                Dim acLine As Line = New Line(PosPoint2, PosPoint1)
                                Dim V3d As Vector3d = PosPoint1.GetVectorTo(PosPoint2)
                                '1.строим 4 точки в начале
                                Dim pt1X As Double = PosPoint1.X + MinHeight * Math.Cos(acLine.Angle + Math.PI / 2)
                                Dim pt1y As Double = PosPoint1.Y + MinHeight * Math.Sin(acLine.Angle + Math.PI / 2)
                                Dim pt1z As Double = PosPoint1.Z + DownHeight
                                Dim PtN1 As Point3d = New Point3d(pt1X, pt1y, pt1z)

                                Dim pt2X As Double = PosPoint1.X + MinHeight * Math.Cos(acLine.Angle - Math.PI / 2)
                                Dim pt2y As Double = PosPoint1.Y + MinHeight * Math.Sin(acLine.Angle - Math.PI / 2)
                                Dim pt2z As Double = PosPoint1.Z + DownHeight
                                Dim PtN2 As Point3d = New Point3d(pt2X, pt2y, pt2z)

                                Dim pt3X As Double = PosPoint1.X + MaxHeight * Math.Cos(acLine.Angle + Math.PI / 2)
                                Dim pt3y As Double = PosPoint1.Y + MaxHeight * Math.Sin(acLine.Angle + Math.PI / 2)
                                Dim pt3z As Double = PosPoint1.Z + TopHeight
                                Dim PtV1 As Point3d = New Point3d(pt3X, pt3y, pt3z)

                                Dim pt4X As Double = PosPoint1.X + MaxHeight * Math.Cos(acLine.Angle - Math.PI / 2)
                                Dim pt4y As Double = PosPoint1.Y + MaxHeight * Math.Sin(acLine.Angle - Math.PI / 2)
                                Dim pt4z As Double = PosPoint1.Z + TopHeight
                                Dim PtV2 As Point3d = New Point3d(pt4X, pt4y, pt4z)

                                Try
                                    'находим точки пересечения с поверхностью
                                    Dim PtV4 As Point3d = TinSurface.GetIntersectionPoint(PtV1, V3d)
                                    Dim PtV3 As Point3d = TinSurface.GetIntersectionPoint(PtV2, V3d)

                                    Dim PtN4 As Point3d = TinSurface.GetIntersectionPoint(PtN1, V3d)
                                    Dim PtN3 As Point3d = TinSurface.GetIntersectionPoint(PtN2, V3d)

                                    Dim Line1 As Line = New Line(PtV1, PtV4)
                                    Dim Line2 As Line = New Line(PtV2, PtV3)

                                    Dim Line3 As Line = New Line(PtN1, PtN4)
                                    Dim Line4 As Line = New Line(PtN2, PtN3)

                                    If Line1.Length > OffsetOtkos Then
                                        PtV4 = Line1.GetPointAtDist(Line1.Length - OffsetOtkos)
                                    End If
                                    If Line2.Length > OffsetOtkos Then
                                        PtV3 = Line2.GetPointAtDist(Line2.Length - OffsetOtkos)
                                    End If
                                    If Line3.Length > OffsetOtkos Then
                                        PtN4 = Line3.GetPointAtDist(Line3.Length - OffsetOtkos)
                                    End If
                                    If Line4.Length > OffsetOtkos Then
                                        PtN3 = Line4.GetPointAtDist(Line4.Length - OffsetOtkos)
                                    End If

                                    '=================================================================
                                    'рисуем 2 замкнутые 3д полилинии
                                    ReDim Preserve ArrayID(2, CountArrayID)
                                    'строим 2 3dполилинии
                                    Dim pts1 As Point3dCollection = New Point3dCollection()
                                    pts1.Add(PtV1)
                                    pts1.Add(PtV2)
                                    pts1.Add(PtV3)
                                    pts1.Add(PtV4)
                                    Using poly1 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts1, True)
                                        acBlkTblRec.AppendEntity(poly1)
                                        acTrans.AddNewlyCreatedDBObject(poly1, True)
                                        ArrayID(1, CountArrayID) = poly1.Id
                                    End Using
                                    Dim PolyPod1 As Polyline3d = Nothing
                                    Dim pts2 As Point3dCollection = New Point3dCollection()
                                    pts2.Add(PtN1)
                                    pts2.Add(PtN2)
                                    pts2.Add(PtN3)
                                    pts2.Add(PtN4)
                                    Using poly2 As Polyline3d = New Polyline3d(Poly3dType.SimplePoly, pts2, True)
                                        acBlkTblRec.AppendEntity(poly2)
                                        acTrans.AddNewlyCreatedDBObject(poly2, True)
                                        ArrayID(0, CountArrayID) = poly2.Id
                                    End Using
                                    Dim LineNapr As Line = New Line(PtV1, PtN1)
                                    acBlkTblRec.AppendEntity(LineNapr)
                                    acTrans.AddNewlyCreatedDBObject(LineNapr, True)
                                    ArrayID(2, CountArrayID) = LineNapr.Id
                                    CountArrayID += 1
                                Catch ex As System.ArgumentException
                                End Try
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            If IsArray(ArrayID) = True Then
                For i As Integer = 0 To ArrayID.GetUpperBound(1)
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ArrayID(0, i), OpenMode.ForRead)
                    Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ArrayID(1, i), OpenMode.ForRead)
                    Dim path As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(ArrayID(2, i), OpenMode.ForRead)
                    Dim options As New LoftOptions
                    Dim sol3d As New Solid3d
                    sol3d.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {ent1, ent2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, Nothing, options)
                    acBlkTblRec.AppendEntity(sol3d)
                    acTrans.AddNewlyCreatedDBObject(sol3d, True)
                Next i
            End If
            acTrans.Commit()
        End Using
    End Sub
    ''=============================================================================================










    ''=============================================================================================
    ''6. Обрезка солида, другим солидом
    ''=============================================================================================
    '<CommandMethod("ADSK-ObrezkaSolid")>
    'Public Sub ObrezkaSolid()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Do
    '        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            'выбираем тело для обрезки
    '            Dim FirstSolid3d As Solid3d = Nothing
    '            Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело которое надо обрезать: ")
    '            If Not (acSSPrompt.Status = PromptStatus.OK) Then
    '                Exit Do
    '            End If
    '            FirstSolid3d = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite)

    '            'выбираем тело по которому будет происходит обрезка
    '            Dim SecondSolid3d As Solid3d = Nothing
    '            Dim acSSPrompt1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело по которому будет происходить обрезка: ")
    '            If Not (acSSPrompt1.Status = PromptStatus.OK) Then
    '                Exit Do
    '            End If
    '            SecondSolid3d = acTrans.GetObject(acSSPrompt1.ObjectId, OpenMode.ForRead)
    '            Dim acSol3DCopy3 As Solid3d = FirstSolid3d.Clone()
    '            If acSol3DCopy3.CheckInterference(SecondSolid3d) = True Then
    '                'acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox3.Clone())
    '                acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolSubtract, SecondSolid3d.Clone())
    '                acSol3DCopy3.ColorIndex = 42
    '                acSol3DCopy3.Layer = FirstSolid3d.Layer
    '                acBlkTblRec.AppendEntity(acSol3DCopy3)
    '                acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)
    '            End If
    '            FirstSolid3d.Erase()
    '            acTrans.Commit()
    '        End Using
    '    Loop
    'End Sub
    ''=============================================================================================
    ''7. Обрезка солида, другим солидом (2 варипнт)
    ''=============================================================================================
    '<CommandMethod("ADSK-ObrezkaSolid2")>
    'Public Sub ObrezkaSolid2()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Do
    '        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '            'выбираем тело для обрезки
    '            Dim FirstSolid3d As Solid3d = Nothing
    '            Dim acSSPrompt As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело которое надо обрезать: ")
    '            If Not (acSSPrompt.Status = PromptStatus.OK) Then
    '                Exit Do
    '            End If
    '            FirstSolid3d = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForWrite)
    '            Dim ext As Extents3d = FirstSolid3d.GeometricExtents
    '            Dim CentrePointX As Double = (ext.MaxPoint.X + ext.MinPoint.X) / 2
    '            Dim CentrePointY As Double = (ext.MaxPoint.Y + ext.MinPoint.Y) / 2
    '            '==============================================================================================
    '            'находим муксимальную и минимальную точку
    '            Dim ptsS As Point3dCollection = New Point3dCollection
    '            Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(FirstSolid3d)
    '                For Each vertex In brep.Vertices
    '                    If Math.Abs(vertex.Point.X - ext.MinPoint.X) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    ElseIf Math.Abs(vertex.Point.X - ext.MaxPoint.X) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    ElseIf Math.Abs(vertex.Point.Y - ext.MaxPoint.Y) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    ElseIf Math.Abs(vertex.Point.Y - ext.MinPoint.Y) < 0.01 Then
    '                        ptsS.Add(vertex.Point)
    '                    End If
    '                Next
    '                If ptsS.Count < 3 Then
    '                    For Each vertex In brep.Vertices
    '                        ptsS.Add(vertex.Point)
    '                    Next
    '                End If
    '            End Using
    '            '===================================================================================
    '            'если в коллекции они есть, то по ним находим максимальный и минимальный пикет
    '            If ptsS.Count > 1 Then
    '                Dim MinH As Double = 99999999
    '                Dim MaxH As Double = 99999999
    '            End If



    '            Dim acSol3DCylindr As Solid3d = New Solid3d()
    '            acSol3DCylindr.SetDatabaseDefaults()
    '            acSol3DCylindr.CreateFrustum(2, 0.4, 0.4, 0.4)
    '            'acSol3DCylindr.TransformBy(Matrix3d.Displacement(New Point3d(CentrePointX, CentrePointY, InsPointStruct.Z - 0.12 + 0.06) - Point3d.Origin))

    '            'выбираем тело по которому будет происходит обрезка
    '            Dim SecondSolid3d As Solid3d = Nothing
    '            Dim acSSPrompt1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите тело по которому будет происходить обрезка: ")
    '            If Not (acSSPrompt1.Status = PromptStatus.OK) Then
    '                Exit Do
    '            End If
    '            SecondSolid3d = acTrans.GetObject(acSSPrompt1.ObjectId, OpenMode.ForRead)
    '            Dim acSol3DCopy3 As Solid3d = FirstSolid3d.Clone()
    '            If acSol3DCopy3.CheckInterference(SecondSolid3d) = True Then
    '                'acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox3.Clone())
    '                acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolSubtract, SecondSolid3d.Clone())
    '                acSol3DCopy3.ColorIndex = 42
    '                acSol3DCopy3.Layer = FirstSolid3d.Layer
    '                acBlkTblRec.AppendEntity(acSol3DCopy3)
    '                acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)
    '            End If
    '            FirstSolid3d.Erase()
    '            acTrans.Commit()
    '        End Using
    '    Loop

    'End Sub
    ''=============================================================================================
    ''8. обрезка песчаной досыпки трубы по колодцу
    ''=============================================================================================
    '<CommandMethod("ADSK-InsertBLKTemp1")>
    'Public Sub InsertBLKTemp1()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim NameObject As String = "Колодец для обрезки10"
    '    Dim PutchFile As String = "d:\Civil3DToolsUtility\FileResources\3DBlock\Structures\Колодцы смотровые подземных коммуникаций.dwg"
    '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        'копируем определение блока
    '        If IsNothing(NameObject) = False Then
    '            If NameObject.Trim.Length > 0 Then
    '                If acBlkTbl.Has(NameObject) = False Then
    '                    Dim boolinsBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFile, NameObject)
    '                End If
    '            End If
    '        End If
    '        'находим колодцы
    '        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '        If acSSPrompt.Status = PromptStatus.OK Then
    '            Dim acSSet As SelectionSet = acSSPrompt.Value
    '            For Each acSSObj As SelectedObject In acSSet
    '                If Not IsDBNull(acSSObj) Then
    '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                    If Not IsDBNull(acEnt) Then
    '                        If TypeOf acEnt Is [Structure] Then
    '                            Dim Struct As [Structure] = acEnt
    '                            Dim structAngle As Double = Struct.Rotation
    '                            Dim PosPoint As Point3d = Struct.Position
    '                            Dim pk As Double = Struct.Station
    '                            Dim PKStr As String = FuncFormatZN.FuncFormatPK(pk, 2)
    '                            'Struct.Erase()
    '                            Dim acBlkObj As BlockReference = AcadFunction.FuncInsertDrawingBlk(PosPoint, NameObject, "0", 7, 1, 20, structAngle, 0)
    '                            'If acBlkObj.Id.IsNull = False Then
    '                            '    Dim ObjColl As DBObjectCollection = New DBObjectCollection
    '                            '    acBlkObj.Explode(ObjColl)
    '                            '    If ObjColl.Count > 0 Then
    '                            '        For Each db As Autodesk.AutoCAD.DatabaseServices.DBObject In ObjColl
    '                            '            Dim Ent As Autodesk.AutoCAD.DatabaseServices.Entity = db
    '                            '            acBlkTblRec.AppendEntity(Ent)
    '                            '            acTrans.AddNewlyCreatedDBObject(Ent, True)
    '                            '            Dim boolnameTbl As Boolean = FuncPropertySet.FuncFindPSetTable(db, "PROJECT_PIPENETWORK")
    '                            '            If boolnameTbl = True Then
    '                            '                Dim ArrayODRec As String(,) = Nothing
    '                            '                ReDim Preserve ArrayODRec(3, 0)
    '                            '                ArrayODRec(0, 0) = "PROJECT_PIPENETWORK"
    '                            '                ArrayODRec(1, 0) = "STATION"
    '                            '                ArrayODRec(2, 0) = PKStr
    '                            '                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(db, ArrayODRec)
    '                            '            End If
    '                            '        Next
    '                            '    End If
    '                            'End If
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '        acTrans.Commit()
    '    End Using
    'End Sub
    ''=============================================================================================
    ''9. Пользовательские свойства трубы
    ''=============================================================================================
    '<CommandMethod("ADSK-AddPSTablePipe")>
    'Public Sub AddPSTablePipe()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim AlignProject As Alignment = Nothing
    '        If CivilDoc.GetAlignmentIds.Count = 1 Then
    '            Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '            AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '        Else
    '            Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent1
    '            End If
    '        End If
    '        If IsNothing(AlignProject) = True Then
    '            Exit Sub
    '        End If
    '        Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_PIPE")
    '        'Выбираем файл Excel
    '        Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '        opFile.InitialDirectory = "c:"
    '        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '        opFile.RestoreDirectory = True
    '        opFile.ShowDialog()
    '        Dim fFile1 As String = opFile.FileName
    '        If IsNothing(fFile1) = True Then Exit Sub
    '        If fFile1.Trim.Length = 0 Then Exit Sub
    '        Dim xlApp As Excel.Application = Nothing
    '        Dim xlBook As Excel.Workbook = Nothing
    '        Dim xlSheets As Excel.Worksheet = Nothing
    '        xlApp = CreateObject("Excel.Application")
    '        xlBook = xlApp.Workbooks.Open(fFile1)
    '        Dim xlsheet As Excel.Worksheet
    '        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '        Dim CountErr As Integer = 100
    '        Dim ArrayDataExcel As Double(,) = Nothing
    '        Dim CountArrayDataExcel As Integer = 0
    '        For indRows As Integer = 6 To xlsheet.Rows.Count - 1
    '            If CountErr = 0 Then Exit For
    '            Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '            If IsNumeric(NumbStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            End If
    '            Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
    '            Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
    '            Dim LPipe As String = xlsheet.Cells(indRows, 4).value  'длина трубы
    '            Dim iPipe As String = xlsheet.Cells(indRows, 5).value 'уклон
    '            Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
    '            Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2
    '            Dim H3 As String = xlsheet.Cells(indRows, 8).value 'H3
    '            Dim H4 As String = xlsheet.Cells(indRows, 9).value 'H4
    '            Dim H5 As String = xlsheet.Cells(indRows, 10).value 'H5
    '            Dim TypeG As String = xlsheet.Cells(indRows, 11).value 'тип гасителя
    '            If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
    '                CountErr -= 1
    '                Continue For
    '            ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            Else
    '                ReDim Preserve ArrayDataExcel(10, CountArrayDataExcel)
    '                If IsNothing(StartLeftPKStr) = False Then
    '                    If StartLeftPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(0, CountArrayDataExcel) = Val(StartLeftPKStr)
    '                        ArrayDataExcel(1, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                If IsNothing(StartRightPKStr) = False Then
    '                    If StartRightPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(1, CountArrayDataExcel) = Val(StartRightPKStr)
    '                        ArrayDataExcel(0, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                ArrayDataExcel(2, CountArrayDataExcel) = Val(LPipe)
    '                ArrayDataExcel(3, CountArrayDataExcel) = Val(iPipe)
    '                ArrayDataExcel(4, CountArrayDataExcel) = Val(H1)
    '                ArrayDataExcel(5, CountArrayDataExcel) = Val(H2)
    '                ArrayDataExcel(6, CountArrayDataExcel) = Val(H3)
    '                ArrayDataExcel(7, CountArrayDataExcel) = Val(H4)
    '                ArrayDataExcel(8, CountArrayDataExcel) = Val(H5)
    '                ArrayDataExcel(9, CountArrayDataExcel) = Val(TypeG)
    '                CountArrayDataExcel += 1
    '                CountErr = 100
    '            End If
    '        Next
    '        xlBook.Close()
    '        xlApp.Quit()
    '        If IsArray(ArrayDataExcel) = True Then
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is Solid3d Then
    '                                Dim acSolid3d As Solid3d = acEnt
    '                                If acSolid3d.Layer Like "Трубы" Then
    '                                    Dim ext As Extents3d = acSolid3d.GeometricExtents
    '                                    Dim Senterpoint As Point2d = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2)
    '                                    Dim PK As Double = -1
    '                                    Dim off As Double = -1
    '                                    Try
    '                                        AlignProject.StationOffset(Senterpoint.X, Senterpoint.Y, PK, off)
    '                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                    End Try
    '                                    If PK <> -1 And off <> -1 Then
    '                                        Dim LenghtPK As Double = 9999999
    '                                        Dim Pos As Integer = 0
    '                                        Dim PKEx As Double = -1
    '                                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                            Dim PKL As Double = ArrayDataExcel(0, i)
    '                                            Dim PKR As Double = ArrayDataExcel(1, i)
    '                                            If PKL <> -777 Then
    '                                                Dim TempL As Double = Math.Abs(PKL - PK)
    '                                                If LenghtPK > TempL Then
    '                                                    LenghtPK = TempL
    '                                                    Pos = i
    '                                                    PKEx = PKL
    '                                                End If
    '                                            Else
    '                                                Dim TempL As Double = Math.Abs(PKR - PK)
    '                                                If LenghtPK > TempL Then
    '                                                    LenghtPK = TempL
    '                                                    Pos = i
    '                                                    PKEx = PKR
    '                                                End If
    '                                            End If
    '                                        Next i
    '                                        'присваиваем расширенные данные
    '                                        Dim PKStr As String = FuncFormatZN.FuncFormatPK(PKEx, 2)
    '                                        Dim ArrayData As String(,) = Nothing
    '                                        ReDim ArrayData(2, 10)
    '                                        ArrayData(0, 0) = "PROJECT_PIPE"
    '                                        ArrayData(1, 0) = "NAME"
    '                                        ArrayData(2, 0) = "Труба двухслойная гофрированная"

    '                                        ArrayData(0, 1) = "PROJECT_PIPE"
    '                                        ArrayData(1, 1) = "GOST"
    '                                        ArrayData(2, 1) = "SN8 DN/ID (300)"

    '                                        ArrayData(0, 2) = "PROJECT_PIPE"
    '                                        ArrayData(1, 2) = "LENGHT"
    '                                        ArrayData(2, 2) = ArrayDataExcel(2, Pos)

    '                                        ArrayData(0, 3) = "PROJECT_PIPE"
    '                                        ArrayData(1, 3) = "END_LEVEL"
    '                                        ArrayData(2, 3) = ArrayDataExcel(6, Pos)

    '                                        ArrayData(0, 4) = "PROJECT_PIPE"
    '                                        ArrayData(1, 4) = "START_LEVEL"
    '                                        ArrayData(2, 4) = ArrayDataExcel(5, Pos)

    '                                        ArrayData(0, 5) = "PROJECT_PIPE"
    '                                        ArrayData(1, 5) = "MATERIAL"
    '                                        ArrayData(2, 5) = "Полиэтилен"

    '                                        ArrayData(0, 6) = "PROJECT_PIPE"
    '                                        ArrayData(1, 6) = "SLOPE"
    '                                        ArrayData(2, 6) = ArrayDataExcel(3, Pos)

    '                                        ArrayData(0, 7) = "PROJECT_PIPE"
    '                                        ArrayData(1, 7) = "STATION"
    '                                        ArrayData(2, 7) = PKStr

    '                                        ArrayData(0, 8) = "PROJECT_PIPE"
    '                                        ArrayData(1, 8) = "DIAMETER_IN"
    '                                        ArrayData(2, 8) = "300"

    '                                        ArrayData(0, 9) = "PROJECT_PIPE"
    '                                        ArrayData(1, 9) = "DIAMETER_OUT"
    '                                        ArrayData(2, 9) = "340"

    '                                        ArrayData(0, 10) = "PROJECT_PIPE"
    '                                        ArrayData(1, 10) = "NOTE"
    '                                        ArrayData(2, 10) = "Тип гасителя - " & ArrayDataExcel(9, Pos)

    '                                        If IsNothing(CreateTBLPSID) = False Then
    '                                            If CreateTBLPSID.IsNull = False Then
    '                                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acSolid3d, CreateTBLPSID)
    '                                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSolid3d, ArrayData)
    '                                            End If
    '                                        End If
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            tr.Commit()
    '        End If
    '    End Using
    'End Sub
    ''=============================================================================================
    ''10. Пользовательские свойства колодца
    ''=============================================================================================
    '<CommandMethod("ADSK-AddPSTableStructure")>
    'Public Sub AddPSTableStructure()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database

    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim AlignProject As Alignment = Nothing
    '        If CivilDoc.GetAlignmentIds.Count = 1 Then
    '            Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '            AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '        Else
    '            Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent1
    '            End If
    '        End If
    '        If IsNothing(AlignProject) = True Then
    '            Exit Sub
    '        End If
    '        Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_STRUCTURE")
    '        'Выбираем файл Excel
    '        Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '        opFile.InitialDirectory = "c:"
    '        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '        opFile.RestoreDirectory = True
    '        opFile.ShowDialog()
    '        Dim fFile1 As String = opFile.FileName
    '        If IsNothing(fFile1) = True Then Exit Sub
    '        If fFile1.Trim.Length = 0 Then Exit Sub
    '        Dim xlApp As Excel.Application = Nothing
    '        Dim xlBook As Excel.Workbook = Nothing
    '        Dim xlSheets As Excel.Worksheet = Nothing
    '        xlApp = CreateObject("Excel.Application")
    '        xlBook = xlApp.Workbooks.Open(fFile1)
    '        Dim xlsheet As Excel.Worksheet
    '        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '        Dim CountErr As Integer = 100
    '        Dim ArrayDataExcel As Double(,) = Nothing
    '        Dim CountArrayDataExcel As Integer = 0
    '        For indRows As Integer = 6 To xlsheet.Rows.Count - 1
    '            If CountErr = 0 Then Exit For
    '            Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '            If IsNumeric(NumbStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            End If
    '            Dim StartLeftPKStr As String = xlsheet.Cells(indRows, 2).value
    '            Dim StartRightPKStr As String = xlsheet.Cells(indRows, 3).value
    '            Dim LPipe As String = xlsheet.Cells(indRows, 4).value  'длина трубы
    '            Dim iPipe As String = xlsheet.Cells(indRows, 5).value 'уклон
    '            Dim H1 As String = xlsheet.Cells(indRows, 6).value 'H1
    '            Dim H2 As String = xlsheet.Cells(indRows, 7).value 'H2
    '            Dim H3 As String = xlsheet.Cells(indRows, 8).value 'H3
    '            Dim H4 As String = xlsheet.Cells(indRows, 9).value 'H4
    '            Dim H5 As String = xlsheet.Cells(indRows, 10).value 'H5
    '            Dim TypeG As String = xlsheet.Cells(indRows, 11).value 'тип гасителя
    '            If IsNothing(StartLeftPKStr) = True And IsNothing(StartRightPKStr) = True Then
    '                CountErr -= 1
    '                Continue For
    '            ElseIf IsNumeric(StartLeftPKStr) = False And IsNumeric(StartRightPKStr) = False Then
    '                CountErr -= 1
    '                Continue For
    '            Else
    '                ReDim Preserve ArrayDataExcel(10, CountArrayDataExcel)
    '                If IsNothing(StartLeftPKStr) = False Then
    '                    If StartLeftPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(0, CountArrayDataExcel) = Val(StartLeftPKStr)
    '                        ArrayDataExcel(1, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                If IsNothing(StartRightPKStr) = False Then
    '                    If StartRightPKStr.Trim.Length > 0 Then
    '                        ArrayDataExcel(1, CountArrayDataExcel) = Val(StartRightPKStr)
    '                        ArrayDataExcel(0, CountArrayDataExcel) = -777
    '                    End If
    '                End If
    '                ArrayDataExcel(2, CountArrayDataExcel) = Val(LPipe)
    '                ArrayDataExcel(3, CountArrayDataExcel) = Val(iPipe)
    '                ArrayDataExcel(4, CountArrayDataExcel) = Val(H1)
    '                ArrayDataExcel(5, CountArrayDataExcel) = Val(H2)
    '                ArrayDataExcel(6, CountArrayDataExcel) = Val(H3)
    '                ArrayDataExcel(7, CountArrayDataExcel) = Val(H4)
    '                ArrayDataExcel(8, CountArrayDataExcel) = Val(H5)
    '                ArrayDataExcel(9, CountArrayDataExcel) = Val(TypeG)
    '                CountArrayDataExcel += 1
    '                CountErr = 100
    '            End If
    '        Next
    '        xlBook.Close()
    '        xlApp.Quit()
    '        If IsArray(ArrayDataExcel) = True Then
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is Solid3d Then
    '                                Dim acSolid3d As Solid3d = acEnt
    '                                Dim boolTbl As Boolean = FuncPropertySet.FuncFindPSetTable(acSolid3d, "PROJECT_STRUCTURE")
    '                                Dim ext As Extents3d = acSolid3d.GeometricExtents
    '                                Dim Senterpoint As Point2d = New Point2d((ext.MaxPoint.X + ext.MinPoint.X) / 2, (ext.MaxPoint.Y + ext.MinPoint.Y) / 2)
    '                                Dim PK As Double = -1
    '                                Dim off As Double = -1
    '                                Try
    '                                    AlignProject.StationOffset(Senterpoint.X, Senterpoint.Y, PK, off)
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                End Try
    '                                If PK <> -1 And off <> -1 Then
    '                                    Dim LenghtPK As Double = 9999999
    '                                    Dim Pos As Integer = 0
    '                                    Dim PKEx As Double = -1
    '                                    For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                        Dim PKL As Double = ArrayDataExcel(0, i)
    '                                        Dim PKR As Double = ArrayDataExcel(1, i)
    '                                        If PKL <> -777 Then
    '                                            Dim TempL As Double = Math.Abs(PKL - PK)
    '                                            If LenghtPK > TempL Then
    '                                                LenghtPK = TempL
    '                                                Pos = i
    '                                                PKEx = PKL
    '                                            End If
    '                                        Else
    '                                            Dim TempL As Double = Math.Abs(PKR - PK)
    '                                            If LenghtPK > TempL Then
    '                                                LenghtPK = TempL
    '                                                Pos = i
    '                                                PKEx = PKR
    '                                            End If
    '                                        End If
    '                                    Next i
    '                                    'находим отметку верха
    '                                    Dim ptsS As Point3dCollection = New Point3dCollection
    '                                    Using brep = New Autodesk.AutoCAD.BoundaryRepresentation.Brep(acSolid3d)
    '                                        For Each vertex In brep.Vertices
    '                                            If Math.Abs(vertex.Point.X - ext.MinPoint.X) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            ElseIf Math.Abs(vertex.Point.X - ext.MaxPoint.X) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            ElseIf Math.Abs(vertex.Point.Y - ext.MaxPoint.Y) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            ElseIf Math.Abs(vertex.Point.Y - ext.MinPoint.Y) < 0.01 Then
    '                                                ptsS.Add(vertex.Point)
    '                                            End If
    '                                        Next
    '                                        If ptsS.Count < 3 Then
    '                                            For Each vertex In brep.Vertices
    '                                                ptsS.Add(vertex.Point)
    '                                            Next
    '                                        End If
    '                                    End Using
    '                                    Dim MaxPk As Double = -99999999
    '                                    If ptsS.Count > 0 Then
    '                                        For Each pt3d As Point3d In ptsS
    '                                            Dim TempH As Double = pt3d.Z
    '                                            If TempH > MaxPk Then
    '                                                MaxPk = TempH
    '                                            End If
    '                                        Next
    '                                    End If
    '                                    'присваиваем расширенные данные
    '                                    Dim PKStr As String = FuncFormatZN.FuncFormatPK(PKEx, 2)
    '                                    Dim NameObj As String = FuncPropertySet.FuncReadValueSetProperties(acSolid3d, "NAME", "PROJECT_STRUCTURE")
    '                                    Dim ArrayData As String(,) = Nothing
    '                                    ReDim ArrayData(2, 1)
    '                                    ArrayData(0, 0) = "PROJECT_STRUCTURE"
    '                                    ArrayData(1, 0) = "TOP_LEVEL"
    '                                    If NameObj.Trim Like "Люк Т (ГОСТ 3634-99)" Then
    '                                        ArrayData(2, 0) = ArrayDataExcel(4, Pos)
    '                                    Else
    '                                        ArrayData(2, 0) = Math.Round(MaxPk, 2)
    '                                    End If
    '                                    ArrayData(0, 1) = "PROJECT_STRUCTURE"
    '                                    ArrayData(1, 1) = "STATION"
    '                                    ArrayData(2, 1) = PKStr

    '                                    If IsNothing(CreateTBLPSID) = False Then
    '                                        If CreateTBLPSID.IsNull = False Then
    '                                            Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSolid3d, ArrayData)
    '                                        End If
    '                                    End If
    '                                End If

    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            tr.Commit()
    '        End If
    '    End Using
    'End Sub
    ''=============================================================================================
    ''трансформирование солида
    ''=============================================================================================
    '<CommandMethod("ADSK-CreateLoft")>
    'Public Sub CreateLoft()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim Db As Database = acDoc.Database
    '    Dim psr As PromptSelectionResult = Ed.GetSelection()
    '    If psr.Status <> PromptStatus.OK Or psr.Value.Count <> 3 Then
    '        Return
    '    End If
    '    Using tr As Transaction = Db.TransactionManager.StartTransaction()
    '        Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(psr.Value(0).ObjectId, OpenMode.ForRead)
    '        Dim ent2 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(psr.Value(1).ObjectId, OpenMode.ForRead)
    '        Dim path As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(psr.Value(2).ObjectId, OpenMode.ForRead)
    '        Dim sld As New Solid3d
    '        Dim options As New LoftOptions
    '        sld.CreateLoftedSolid(New Autodesk.AutoCAD.DatabaseServices.Entity() {ent1, ent2}, New Autodesk.AutoCAD.DatabaseServices.Entity() {}, path, options)
    '        Dim bt As BlockTable = tr.GetObject(Db.BlockTableId, OpenMode.ForRead)
    '        Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        ms.AppendEntity(sld)
    '        tr.AddNewlyCreatedDBObject(sld, True)
    '        tr.Commit()
    '    End Using
    'End Sub
    ''=============================================================================================
    ''9. Пользовательские свойства открытого водоотвода
    ''=============================================================================================
    '<CommandMethod("ADSK-AddPSVOD")>
    'Public Sub AddPSVOD()
    '    Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    '    Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
    '    Dim id As ObjectIdCollection = CivilDoc.GetNoteLabelIds
    '    Dim Ed As Editor = docCol.MdiActiveDocument.Editor
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
    '    Using tr
    '        Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        Dim AlignProject As Alignment = Nothing
    '        If CivilDoc.GetAlignmentIds.Count = 1 Then
    '            Dim IdAlign As ObjectId = CivilDoc.GetAlignmentIds.Item(0)
    '            AlignProject = tr.GetObject(IdAlign, OpenMode.ForRead)
    '        Else
    '            Dim ppr1 As PromptEntityResult = Ed.GetEntity(vbLf & "Выберите трассу автодороги: ")
    '            Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
    '            If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
    '                AlignProject = ent1
    '            End If
    '        End If
    '        If IsNothing(AlignProject) = True Then
    '            Exit Sub
    '        End If
    '        Dim CreateTBLPSID As ObjectId = FuncPropertySet.FuncReturnIDDictSetProperties("PROJECT_OBJECT")
    '        'Выбираем файл Excel
    '        Dim opFile As New System.Windows.Forms.OpenFileDialog()
    '        opFile.InitialDirectory = "d:\"
    '        opFile.Filter = "Excel files (*.xlsx)|*.xlsx"
    '        opFile.RestoreDirectory = True
    '        opFile.ShowDialog()
    '        Dim fFile1 As String = opFile.FileName
    '        If IsNothing(fFile1) = True Then Exit Sub
    '        If fFile1.Trim.Length = 0 Then Exit Sub
    '        Dim xlApp As Excel.Application = Nothing
    '        Dim xlBook As Excel.Workbook = Nothing
    '        Dim xlSheets As Excel.Worksheet = Nothing
    '        xlApp = CreateObject("Excel.Application")
    '        xlBook = xlApp.Workbooks.Open(fFile1)
    '        Dim xlsheet As Excel.Worksheet
    '        xlsheet = xlBook.Worksheets.Item(1) 'делаем ссылку на этот лист
    '        Dim CountColumn As Integer = -1
    '        For i As Integer = 1 To xlsheet.Columns.Count - 1
    '            Dim TempStr As String = xlsheet.Cells(1, i).value 'H1
    '            If IsNothing(TempStr) = True Then
    '                Exit For
    '            Else
    '                CountColumn += 1
    '            End If
    '        Next i
    '        If CountColumn < 0 Then
    '            MsgBox("Не удалось прчитать наименования столбцов. Отформатируйте таблицу!!!")
    '            Exit Sub
    '        End If
    '        Dim ArrayDataExcel As String(,) = Nothing
    '        Dim CountArrayDataExcel As Integer = 2
    '        ReDim Preserve ArrayDataExcel(CountColumn, CountArrayDataExcel)
    '        '====================================================================================
    '        'читаем наименования столбцов
    '        For i As Integer = 1 To xlsheet.Columns.Count - 1
    '            Dim TempStr As String = xlsheet.Cells(1, i).value 'Наименование столбца
    '            Dim TempStr2 As String = xlsheet.Cells(2, i).value 'Еденица измерения
    '            If IsNothing(TempStr) = True Then
    '                Exit For
    '            Else
    '                If i - 1 < CountColumn Then
    '                    ArrayDataExcel(i - 1, 0) = TempStr
    '                    ArrayDataExcel(i - 1, 1) = TempStr2
    '                End If
    '            End If
    '        Next i
    '        '==================================================================================
    '        'читаем строки
    '        For indRows As Integer = 3 To xlsheet.Rows.Count - 1
    '            Dim NumbStr As String = xlsheet.Cells(indRows, 1).value
    '            If IsNothing(NumbStr) = True Then
    '                Exit For
    '            End If
    '            ReDim Preserve ArrayDataExcel(CountColumn, CountArrayDataExcel)
    '            For indCol As Integer = 1 To ArrayDataExcel.GetUpperBound(0)
    '                ArrayDataExcel(indCol - 1, CountArrayDataExcel) = xlsheet.Cells(indRows, indCol).value
    '            Next
    '            CountArrayDataExcel += 1
    '        Next
    '        xlBook.Close()
    '        xlApp.Quit()
    '        If IsArray(ArrayDataExcel) = True Then
    '            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '            If acSSPrompt.Status = PromptStatus.OK Then
    '                Dim acSSet As SelectionSet = acSSPrompt.Value
    '                For Each acSSObj As SelectedObject In acSSet
    '                    If Not IsDBNull(acSSObj) Then
    '                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                        If Not IsDBNull(acEnt) Then
    '                            If TypeOf acEnt Is Solid3d Then
    '                                Dim acSolid3d As Solid3d = acEnt
    '                                Dim LayerSolid As String = acSolid3d.Layer
    '                                Dim Ext As Extents3d = acSolid3d.Bounds
    '                                Dim CenterPoint As Point2d = New Point2d((Ext.MaxPoint.X + Ext.MinPoint.X) / 2, (Ext.MaxPoint.Y + Ext.MinPoint.Y) / 2)
    '                                Dim PK As Double = -1
    '                                Dim off As Double = -1
    '                                Dim PKFindExcel As Double = -1
    '                                Try
    '                                    AlignProject.StationOffset(CenterPoint.X, CenterPoint.Y, PK, off)
    '                                    Dim PosStroka As Integer = -1
    '                                    Dim MaxLenght As Double = 9999999
    '                                    If off < 0 Then
    '                                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                            Dim PKLStr As String = ArrayDataExcel(1, i)
    '                                            If IsNothing(PKLStr) = False Then
    '                                                Dim PKL As Double = Val(PKLStr)
    '                                                Dim TempLenght As Double = Math.Abs(PKL - PK)
    '                                                If TempLenght < MaxLenght Then
    '                                                    MaxLenght = TempLenght
    '                                                    PosStroka = i
    '                                                    PKFindExcel = PKL
    '                                                End If
    '                                            End If
    '                                        Next
    '                                    Else
    '                                        For i As Integer = 0 To ArrayDataExcel.GetUpperBound(1)
    '                                            Dim PKLStr As String = ArrayDataExcel(2, i)
    '                                            If IsNothing(PKLStr) = False Then
    '                                                Dim PKL As Double = Val(PKLStr)
    '                                                Dim TempLenght As Double = Math.Abs(PKL - PK)
    '                                                If TempLenght < MaxLenght Then
    '                                                    MaxLenght = TempLenght
    '                                                    PosStroka = i
    '                                                    PKFindExcel = PKL
    '                                                End If
    '                                            End If
    '                                        Next
    '                                    End If
    '                                    'ищем столбец
    '                                    Dim StrWrite As String = ""
    '                                    Dim EdFormat As String = ""
    '                                    For i As Integer = 0 To ArrayDataExcel.GetUpperBound(0)
    '                                        Dim TempField As String = ArrayDataExcel(i, 0)
    '                                        If IsNothing(TempField) = False Then
    '                                            If TempField.Trim Like LayerSolid Then
    '                                                StrWrite = ArrayDataExcel(i, PosStroka)
    '                                                EdFormat = ArrayDataExcel(i, 1)
    '                                                'присваиваем расширенные данные
    '                                                If IsNothing(CreateTBLPSID) = False Then
    '                                                    If CreateTBLPSID.IsNull = False Then
    '                                                        Dim PKStr As String = FuncFormatZN.FuncFormatPK(PKFindExcel, 2)
    '                                                        Dim ArrayData As String(,) = Nothing
    '                                                        ReDim ArrayData(2, 2)

    '                                                        ArrayData(0, 0) = "PROJECT_OBJECT"
    '                                                        ArrayData(1, 0) = "NAME"
    '                                                        ArrayData(2, 0) = LayerSolid

    '                                                        ArrayData(0, 1) = "PROJECT_OBJECT"
    '                                                        ArrayData(1, 1) = "STATION"
    '                                                        ArrayData(2, 1) = PKStr

    '                                                        If EdFormat Like "шт" Then
    '                                                            ArrayData(0, 2) = "PROJECT_OBJECT"
    '                                                            ArrayData(1, 2) = "VALUE1"
    '                                                            ArrayData(2, 2) = StrWrite
    '                                                        ElseIf EdFormat Like "м2" Then
    '                                                            ArrayData(0, 2) = "PROJECT_OBJECT"
    '                                                            ArrayData(1, 2) = "VALUE2"
    '                                                            ArrayData(2, 2) = StrWrite
    '                                                        Else
    '                                                            ArrayData(0, 2) = "PROJECT_OBJECT"
    '                                                            ArrayData(1, 2) = "VALUE3"
    '                                                            ArrayData(2, 2) = StrWrite
    '                                                        End If
    '                                                        If IsNothing(CreateTBLPSID) = False Then
    '                                                            If CreateTBLPSID.IsNull = False Then
    '                                                                Dim boolAddTLB As Boolean = FuncPropertySet.FuncAddDataSetProperties(acSolid3d, CreateTBLPSID)
    '                                                                Dim boolWriteData As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acSolid3d, ArrayData)
    '                                                            End If
    '                                                        End If
    '                                                    End If
    '                                                End If

    '                                            End If
    '                                        Else
    '                                            Continue For
    '                                        End If
    '                                    Next
    '                                Catch ex As Autodesk.Civil.PointNotOnEntityException
    '                                End Try
    '                            End If
    '                        End If
    '                    End If
    '                Next
    '            End If
    '            tr.Commit()
    '        End If
    '    End Using
    'End Sub

    '<CommandMethod("ExportPipeNetworkToSolid3d")>
    'Public Sub ExportPipeNetworkToSolid3d()
    '    '' Get the current document and database, and start a transaction
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim m_Editor As Editor = acDoc.Editor
    '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        'создаем нове слои
    '        '1. Решетки сточные
    '        Dim NameLayerRS As String = "Решетка сточная"
    '        Dim NameLayerPP As String = "Плита перекрытия"
    '        Dim NameLayerK1 As String = "Стеновое кольцо1"
    '        Dim NameLayerK2 As String = "Стеновое Кольцо2"
    '        Dim NameLayerD As String = "Плита днища"
    '        Dim NameLayerPipe As String = "Трубы"
    '        Dim NameLayerPS As String = "Портальная стенка"
    '        Dim idL1 As ObjectId = AcadFunction.FuncAddLayer(NameLayerRS, 7,, 20)
    '        Dim idL2 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPP, 7,, 20)
    '        Dim idL3 As ObjectId = AcadFunction.FuncAddLayer(NameLayerK1, 7,, 20)
    '        Dim idL4 As ObjectId = AcadFunction.FuncAddLayer(NameLayerK2, 7,, 20)
    '        Dim idL5 As ObjectId = AcadFunction.FuncAddLayer(NameLayerD, 7,, 20)
    '        Dim idL6 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPipe, 7,, 20)
    '        Dim idL7 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPS, 7,, 20)
    '        Dim curUCSMatrix As Matrix3d = acDoc.Editor.CurrentUserCoordinateSystem
    '        Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
    '        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '        If acSSPrompt.Status = PromptStatus.OK Then
    '            Dim acSSet As SelectionSet = acSSPrompt.Value
    '            For Each acSSObj As SelectedObject In acSSet
    '                If Not IsDBNull(acSSObj) Then
    '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                    If Not IsDBNull(acEnt) Then
    '                        If TypeOf acEnt Is [Structure] Then
    '                            Dim Struct As [Structure] = acEnt
    '                            Dim structAngle As Double = Struct.Rotation
    '                            'преобразуем в солид
    '                            Dim acSolid3d As Solid3d = Struct.Solid3dBody
    '                            'находим самую верхнюю точку
    '                            Dim InsPointStruct As Point3d = Struct.Position
    '                            'строим вспомогательные солиды для разрезки колодца
    '                            Dim acSol3DBox1 As Solid3d = New Solid3d()
    '                            acSol3DBox1.SetDatabaseDefaults()
    '                            acSol3DBox1.CreateBox(2, 2, 0.12)
    '                            acSol3DBox1.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 0.12 + 0.06) - Point3d.Origin))

    '                            Dim acSol3DBox2 As Solid3d = New Solid3d()
    '                            acSol3DBox2.SetDatabaseDefaults()
    '                            acSol3DBox2.CreateBox(2, 2, 0.16)
    '                            acSol3DBox2.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 0.28 + 0.08) - Point3d.Origin))
    '                            acSol3DBox2.TransformBy(Matrix3d.Rotation(structAngle, curUCS.Zaxis, Struct.Position))

    '                            'вспомогательный солид для сточной решетки
    '                            Dim acSol3DBox_RS As Solid3d = New Solid3d()
    '                            acSol3DBox_RS.SetDatabaseDefaults()
    '                            acSol3DBox_RS.CreateBox(0.4, 0.8, 0.16)
    '                            acSol3DBox_RS.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 0.28 + 0.08) - Point3d.Origin))
    '                            acSol3DBox_RS.TransformBy(Matrix3d.Rotation(structAngle, curUCS.Zaxis, Struct.Position))

    '                            'вспомогательный цилиндр для выдавливания кольца
    '                            Dim acSol3DCyl As Solid3d = New Solid3d()
    '                            acSol3DCyl.SetDatabaseDefaults()
    '                            acSol3DCyl.CreateFrustum(2, 0.4, 0.4, 0.4)
    '                            acSol3DCyl.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1) - Point3d.Origin))

    '                            Dim acSol3DBox3 As Solid3d = New Solid3d()
    '                            acSol3DBox3.SetDatabaseDefaults()
    '                            acSol3DBox3.CreateBox(2, 2, 0.765)
    '                            acSol3DBox3.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1.045 + 0.3825) - Point3d.Origin))

    '                            Dim acSol3DBox4 As Solid3d = New Solid3d()
    '                            acSol3DBox4.SetDatabaseDefaults()
    '                            acSol3DBox4.CreateBox(2, 2, 0.765)
    '                            acSol3DBox4.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1.81 + 0.3825) - Point3d.Origin))

    '                            Dim acSol3DBox5 As Solid3d = New Solid3d()
    '                            acSol3DBox5.SetDatabaseDefaults()
    '                            acSol3DBox5.CreateBox(2, 2, 0.1)
    '                            acSol3DBox5.TransformBy(Matrix3d.Displacement(New Point3d(InsPointStruct.X, InsPointStruct.Y, InsPointStruct.Z - 1.91 + 0.05) - Point3d.Origin))

    '                            'разрезаем фигуры
    '                            Dim acSol3DCopy1 As Solid3d = acSolid3d.Clone()
    '                            If acSol3DCopy1.CheckInterference(acSol3DBox1) = True Then
    '                                acSol3DCopy1.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox1.Clone())
    '                                acSol3DCopy1.ColorIndex = 1
    '                                acSol3DCopy1.Layer = NameLayerRS
    '                                acBlkTblRec.AppendEntity(acSol3DCopy1)
    '                                acTrans.AddNewlyCreatedDBObject(acSol3DCopy1, True)
    '                            End If
    '                            Dim acSol3DCopy2 As Solid3d = acSolid3d.Clone()
    '                            If acSol3DCopy2.CheckInterference(acSol3DBox2) = True Then
    '                                acSol3DCopy2.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox2.Clone())
    '                                acSol3DCopy2.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DBox_RS.Clone())
    '                                acSol3DCopy2.ColorIndex = 2
    '                                acSol3DCopy2.Layer = NameLayerPP
    '                                acBlkTblRec.AppendEntity(acSol3DCopy2)
    '                                acTrans.AddNewlyCreatedDBObject(acSol3DCopy2, True)
    '                            End If
    '                            Dim acSol3DCopy3 As Solid3d = acSolid3d.Clone()
    '                            If acSol3DCopy3.CheckInterference(acSol3DBox3) = True Then
    '                                acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox3.Clone())
    '                                acSol3DCopy3.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DCyl.Clone())
    '                                acSol3DCopy3.ColorIndex = 3
    '                                acSol3DCopy3.Layer = NameLayerK1
    '                                acBlkTblRec.AppendEntity(acSol3DCopy3)
    '                                acTrans.AddNewlyCreatedDBObject(acSol3DCopy3, True)
    '                            End If
    '                            Dim acSol3DCopy4 As Solid3d = acSolid3d.Clone()
    '                            If acSol3DCopy4.CheckInterference(acSol3DBox4) = True Then
    '                                acSol3DCopy4.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox4.Clone())
    '                                acSol3DCopy4.BooleanOperation(BooleanOperationType.BoolSubtract, acSol3DCyl.Clone())
    '                                acSol3DCopy4.ColorIndex = 4
    '                                acSol3DCopy4.Layer = NameLayerK2
    '                                acBlkTblRec.AppendEntity(acSol3DCopy4)
    '                                acTrans.AddNewlyCreatedDBObject(acSol3DCopy4, True)
    '                            End If
    '                            Dim acSol3DCopy5 As Solid3d = acSolid3d.Clone()
    '                            If acSol3DCopy5.CheckInterference(acSol3DBox5) = True Then
    '                                acSol3DCopy5.BooleanOperation(BooleanOperationType.BoolIntersect, acSol3DBox5.Clone())
    '                                acSol3DCopy5.ColorIndex = 5
    '                                acSol3DCopy5.Layer = NameLayerD
    '                                acBlkTblRec.AppendEntity(acSol3DCopy5)
    '                                acTrans.AddNewlyCreatedDBObject(acSol3DCopy5, True)
    '                            End If
    '                        ElseIf TypeOf acEnt Is Pipe Then
    '                            Dim acPipe As Pipe = acEnt
    '                            Dim acSolidPipe As Solid3d = acPipe.Solid3dBody
    '                            acSolidPipe.Layer = NameLayerPipe
    '                            Dim acLine As Line = New Line(acPipe.StartPoint, acPipe.EndPoint)
    '                            Dim PipeAngle As Double = acLine.Angle
    '                            Dim PtPl As Point3d = acLine.GetPointAtDist(acLine.Length - 1)
    '                            Dim acSol3DBox4 As Solid3d = New Solid3d()
    '                            acSol3DBox4.SetDatabaseDefaults()
    '                            acSol3DBox4.CreateBox(0.3, 0.5, 1)
    '                            acSol3DBox4.TransformBy(Matrix3d.Displacement(New Point3d(PtPl.X, PtPl.Y, PtPl.Z + 0.15) - Point3d.Origin))
    '                            acSol3DBox4.TransformBy(Matrix3d.Rotation(PipeAngle, curUCS.Zaxis, PtPl))
    '                            acSol3DBox4.Layer = NameLayerPS
    '                            acBlkTblRec.AppendEntity(acSol3DBox4)
    '                            acTrans.AddNewlyCreatedDBObject(acSol3DBox4, True)
    '                            acBlkTblRec.AppendEntity(acSolidPipe)
    '                            acTrans.AddNewlyCreatedDBObject(acSolidPipe, True)
    '                            acSol3DBox4.BooleanOperation(BooleanOperationType.BoolSubtract, acSolidPipe.Clone())
    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '        acTrans.Commit()
    '    End Using
    'End Sub

    '<CommandMethod("ExportPipeNetworkToSolid3d2")>
    'Public Sub ExportPipeNetworkToSolid3d2()
    '    '' Get the current document and database, and start a transaction
    '    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim acCurDb As Database = acDoc.Database
    '    Dim m_Editor As Editor = acDoc.Editor
    '    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
    '        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
    '        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
    '        'создаем нове слои
    '        '1. Решетки сточные
    '        Dim NameLayerPP As String = "Песчаная подготовка"
    '        Dim NameLayerSH As String = "Щебеночная подготовка под портальную стенку"
    '        Dim NameLayerPipe As String = "Трубы"
    '        Dim NameLayerPS As String = "Портальная стенка"
    '        Dim boolL4 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPP, 7,, 20)
    '        Dim boolL5 As ObjectId = AcadFunction.FuncAddLayer(NameLayerSH, 7,, 20)
    '        Dim boolL6 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPipe, 7,, 20)
    '        Dim boolL7 As ObjectId = AcadFunction.FuncAddLayer(NameLayerPS, 7,, 20)
    '        Dim curUCSMatrix As Matrix3d = acDoc.Editor.CurrentUserCoordinateSystem
    '        Dim curUCS As CoordinateSystem3d = curUCSMatrix.CoordinateSystem3d
    '        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
    '        If acSSPrompt.Status = PromptStatus.OK Then
    '            Dim acSSet As SelectionSet = acSSPrompt.Value
    '            For Each acSSObj As SelectedObject In acSSet
    '                If Not IsDBNull(acSSObj) Then
    '                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
    '                    If Not IsDBNull(acEnt) Then
    '                        If TypeOf acEnt Is Pipe Then
    '                            Dim acPipe As Pipe = acEnt
    '                            Dim acSolidPipe As Solid3d = acPipe.Solid3dBody
    '                            acBlkTblRec.AppendEntity(acSolidPipe)
    '                            acTrans.AddNewlyCreatedDBObject(acSolidPipe, True)
    '                            acSolidPipe.Layer = NameLayerPipe
    '                            Dim acLine As Line = New Line(acPipe.StartPoint, acPipe.EndPoint)
    '                            Dim PipeAngle As Double = acLine.Angle
    '                            Dim PtPl As Point3d = acLine.GetPointAtDist(acLine.Length - 1)
    '                            'строим портальную стенку
    '                            Dim acSol3DBox4 As Solid3d = New Solid3d()
    '                            acSol3DBox4.SetDatabaseDefaults()
    '                            acSol3DBox4.CreateBox(0.3, 1.2, 1)
    '                            acSol3DBox4.TransformBy(Matrix3d.Displacement(New Point3d(PtPl.X, PtPl.Y, PtPl.Z + 0.15) - Point3d.Origin))
    '                            acSol3DBox4.TransformBy(Matrix3d.Rotation(PipeAngle, curUCS.Zaxis, PtPl))
    '                            acSol3DBox4.Layer = NameLayerPS
    '                            acBlkTblRec.AppendEntity(acSol3DBox4)
    '                            acTrans.AddNewlyCreatedDBObject(acSol3DBox4, True)
    '                            acBlkTblRec.AppendEntity(acSolidPipe)
    '                            acTrans.AddNewlyCreatedDBObject(acSolidPipe, True)
    '                            acSol3DBox4.BooleanOperation(BooleanOperationType.BoolSubtract, acSolidPipe.Clone())
    '                            'строим песчаную подготовку
    '                            Dim acSol3DBox5 As Solid3d = New Solid3d()
    '                            acSol3DBox5.SetDatabaseDefaults()
    '                            acSol3DBox5.CreateBox(0.4, 1.2, 0.1)
    '                            acSol3DBox5.TransformBy(Matrix3d.Displacement(New Point3d(PtPl.X, PtPl.Y, PtPl.Z + 0.15) - Point3d.Origin))
    '                            acSol3DBox5.TransformBy(Matrix3d.Rotation(PipeAngle, curUCS.Zaxis, PtPl))
    '                            acSol3DBox5.Layer = NameLayerPP
    '                            acBlkTblRec.AppendEntity(acSol3DBox5)
    '                            acTrans.AddNewlyCreatedDBObject(acSol3DBox5, True)
    '                            'давим трубу
    '                            Dim acSol3DBox6 As Solid3d = New Solid3d()
    '                            acSol3DBox6.SetDatabaseDefaults()

    '                        End If
    '                    End If
    '                End If
    '            Next
    '        End If
    '        acTrans.Commit()
    '    End Using
    'End Sub

    '<CommandMethod("SAP")>
    'Public Sub SweepAlongPath()
    '    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database
    '    Dim ed As Editor = doc.Editor
    '    Dim peo1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select profile or curve to sweep: ")
    '    peo1.SetRejectMessage(vbLf & "Entity must be a region, curve or planar surface.")
    '    peo1.AddAllowedClass(GetType(Region), False)
    '    peo1.AddAllowedClass(GetType(Curve), False)
    '    peo1.AddAllowedClass(GetType(PlaneSurface), False)
    '    Dim per As PromptEntityResult = ed.GetEntity(peo1)
    '    If per.Status <> PromptStatus.OK Then Return
    '    Dim regId As ObjectId = per.ObjectId
    '    Dim peo2 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Select path along which to sweep: ")
    '    peo2.SetRejectMessage(vbLf & "Entity must be a curve.")
    '    peo2.AddAllowedClass(GetType(Curve), False)
    '    per = ed.GetEntity(peo2)
    '    If per.Status <> PromptStatus.OK Then Return
    '    Dim splId As ObjectId = per.ObjectId
    '    Dim pko As PromptKeywordOptions = New PromptKeywordOptions(vbLf & "Sweep a solid or a surface?")
    '    pko.AllowNone = True
    '    pko.Keywords.Add("SOlid")
    '    pko.Keywords.Add("SUrface")
    '    pko.Keywords.[Default] = "SOlid"
    '    Dim pkr As PromptResult = ed.GetKeywords(pko)
    '    Dim createSolid As Boolean = (pkr.StringResult = "SOlid")
    '    If pkr.Status <> PromptStatus.OK Then Return
    '    Dim tr As Transaction = db.TransactionManager.StartTransaction()

    '    Using tr

    '        Try
    '            Dim sweepEnt As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(tr.GetObject(regId, OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
    '            Dim pathEnt As Curve = TryCast(tr.GetObject(splId, OpenMode.ForRead), Curve)

    '            If sweepEnt Is Nothing OrElse pathEnt Is Nothing Then
    '                ed.WriteMessage(vbLf & "Problem opening the selected entities.")
    '                Return
    '            End If

    '            Dim sob As SweepOptionsBuilder = New SweepOptionsBuilder()
    '            sob.Align = SweepOptionsAlignOption.AlignSweepEntityToPath
    '            sob.BasePoint = pathEnt.StartPoint
    '            sob.Bank = True
    '            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity

    '            If createSolid Then
    '                Dim sol As Solid3d = New Solid3d()
    '                sol.CreateSweptSolid(sweepEnt, pathEnt, sob.ToSweepOptions())
    '                ent = sol
    '            Else
    '                Dim ss As SweptSurface = New SweptSurface()
    '                ss.CreateSweptSurface(sweepEnt, pathEnt, sob.ToSweepOptions())
    '                ent = ss
    '            End If

    '            Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
    '            Dim ms As BlockTableRecord = CType(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
    '            ms.AppendEntity(ent)
    '            tr.AddNewlyCreatedDBObject(ent, True)
    '            tr.Commit()
    '        Catch
    '        End Try
    '    End Using
    'End Sub

End Class
