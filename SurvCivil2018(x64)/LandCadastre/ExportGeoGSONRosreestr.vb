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
Public Class ExportGeoGSONRosreestr
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'загрузка объектов в geojson для росреестра
    <CommandMethod("ExportObjectToGeoGSON")>
    Public Sub ExportOBJtoGeoGSON()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim strDWGName As String = acDoc.Name
        Dim NameDwg As String = Path.GetFileNameWithoutExtension(strDWGName)
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
        Dim roundZn As Integer = 8
        '====================================================================================================
        'обозначение участка
        Dim pKeyOpts As PromptKeywordOptions = New PromptKeywordOptions("")
        pKeyOpts.Message = vbLf & "Система координат: "
        pKeyOpts.Keywords.Add("Wgs84")
        pKeyOpts.Keywords.Add("Pseudomercator")
        pKeyOpts.Keywords.Add("Msk")
        pKeyOpts.AllowNone = False
        Dim pKeyRes As PromptResult = acDoc.Editor.GetKeywords(pKeyOpts)
        Dim fileSaveDLG As SaveFileDialog = New SaveFileDialog()
        fileSaveDLG.Filter = ".geojson Files (*.geojson)|*.geojson"
        If fileSaveDLG.ShowDialog() = DialogResult.OK Then
            Dim myStream As Stream = fileSaveDLG.OpenFile()
            Using Tr As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = Tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = Tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                '1.считывае координатную зону чертежа
                Try
                    Dim CoordSystemCode As String = CivilApplication.ActiveDocument.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
                    Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
                    Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
                    Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
                    Dim datumDict As MgCoordinateSystemDatumDictionary = csCatalog.GetDatumDictionary()
                    Dim ellipsoidDict As MgCoordinateSystemEllipsoidDictionary = csCatalog.GetEllipsoidDictionary
                    Dim DrawingSysCoord As MgCoordinateSystem = csDict.GetCoordinateSystem(CoordSystemCode)
                    Try
                        DrawingSysCoord = csDict.GetCoordinateSystem(CoordSystemCode)
                    Catch ex As OSGeo.MapGuide.MgCoordinateSystemLoadFailedException
                        MsgBox("Чертежу не назначена система координат!!! Макро прервано!!!")
                        Exit Sub
                    End Try
                    '2. Пытаемся найти целевую СК
                    Dim coordSystem As MgCoordinateSystem = Nothing
                    If pKeyRes.StringResult Like "Wgs84" Then
                        coordSystem = csDict.GetCoordinateSystem("LL84")
                    ElseIf pKeyRes.StringResult Like "Pseudomercator" Then
                        coordSystem = csDict.GetCoordinateSystem("WGS84.PseudoMercator")
                    Else
                        coordSystem = csDict.GetCoordinateSystem(CoordSystemCode)
                    End If
                    If IsNothing(CoordSystemCode) = False Then
                        If coordSystem.Units Like "DEGREE" Then
                            roundZn = 8
                        ElseIf coordSystem.Units Like "Meter" Then
                            roundZn = 3
                        End If
                        '3. Открываем текстовый файл на запись
                        Using writer As StreamWriter = New StreamWriter(myStream)
                            writer.WriteLine("{")
                            writer.WriteLine("  " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "FeatureCollection" & Chr(34) & ",")
                            writer.WriteLine("  " & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & NameDwg & Chr(34) & ",")
                            writer.WriteLine("  " & Chr(34) & "crs" & Chr(34) & ": { " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "name" & Chr(34) & ", " & Chr(34) & "properties" & Chr(34) & ": {" & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & "urn:ogc:def:crs:EPSG::" & coordSystem.CsCode & Chr(34) & " } },")
                            writer.WriteLine("  " & Chr(34) & "features" & Chr(34) & ": [")
                            writer.WriteLine("    {")
                            writer.WriteLine("      " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Feature" & Chr(34) & ",")
                            writer.WriteLine("      " & Chr(34) & "geometry" & Chr(34) & ": {")
                            '3. Выбираем контуры
                            Dim ObjIDCollections As ObjectIdCollection = New ObjectIdCollection
                            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                            If acSSPrompt.Status = PromptStatus.OK Then
                                Dim acSSet As SelectionSet = acSSPrompt.Value
                                For Each acSSObj As SelectedObject In acSSet
                                    If Not IsDBNull(acSSObj) Then
                                        Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acSSObj.ObjectId.GetObject(OpenMode.ForRead)
                                        Dim arrayTable As String() = {}
                                        If TypeOf ent Is Polyline Then
                                            Dim boolPSTbl As Boolean = FuncPropertySet.FuncReadAllDataSetTables(ent, arrayTable)
                                            Dim boolTable As Boolean = MathFunction.FuncFindValueToArray(arrayTable, "CADASTRAL_OSTROV")
                                            If boolTable = True Then
                                                If arrayTable.Length = 1 Then
                                                    Continue For
                                                End If
                                            End If
                                            ObjIDCollections.Add(acSSObj.ObjectId)
                                        End If
                                    End If
                                Next
                            End If

                            Dim count As Integer = 0
                            If ObjIDCollections.Count > 0 Then
                                writer.WriteLine("        " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "MultiPolygon" & Chr(34) & ",")
                                writer.WriteLine("        " & Chr(34) & "coordinates" & Chr(34) & ": [")
                                For Each id As ObjectId In ObjIDCollections
                                    Dim ind As Integer = ObjIDCollections.IndexOf(id)
                                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = Tr.GetObject(id, OpenMode.ForWrite, False) 'открываем объект на чтение
                                    Dim Pt3dColl As Point3dCollection = New Point3dCollection
                                    If TypeOf ent Is Polyline Then
                                        writer.WriteLine("          [")
                                        writer.WriteLine("            [")
                                        Dim acPoly As Polyline = ent
                                        Dim boolVertRed As Boolean = AcadFunction.FuncPlineDelDubleVertex(acPoly, 0.05)
                                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                        'ищем островки
                                        Dim localID As String = FuncPropertySet.FuncReadValueSetProperties(acPoly, "LOCALID")
                                        Dim arrayOstrov As String(,) = {}
                                        If localID.Trim.Length > 0 Then
                                            Dim boolFindOstrov As Boolean = FuncCadastr.FuncFindOstrov(localID, arrayOstrov)
                                        End If
                                        Dim startPoint As Double = 0
                                        Dim endPoint As Double = 0
                                        '===================================================================================================================================
                                        'записываем координаты
                                        Dim ptPrev As Point2d = New Point2d(-1, -1)
                                        For i As Integer = 0 To acPoly.NumberOfVertices - 1
                                            Dim pt1 As Point2d = acPoly.GetPoint2dAt(i)
                                            If pt1.GetVectorTo(ptPrev).Length < 0.05 Then
                                                Continue For
                                            Else
                                                ptPrev = pt1
                                            End If
                                            Dim x As Double = pt1.X
                                            Dim y As Double = pt1.Y
                                            Dim x_Pseudo As Double = x
                                            Dim y_Pseudo As Double = y
                                            If IsNothing(coordSystem) = False Then
                                                Dim UserCoord_to_PseudoMarcator As MgCoordinateSystemTransform = coordSysFactory.GetTransform(DrawingSysCoord, coordSystem)
                                                Dim UserCoord_to_PseudoMercatorRez As MgCoordinate = UserCoord_to_PseudoMarcator.Transform(x, y)
                                                x_Pseudo = UserCoord_to_PseudoMercatorRez.X
                                                y_Pseudo = UserCoord_to_PseudoMercatorRez.Y
                                            End If
                                            If i = 0 Then
                                                startPoint = x_Pseudo
                                                endPoint = y_Pseudo
                                            End If
                                            writer.WriteLine("              [" & Math.Round(x_Pseudo, roundZn) & ", " & Math.Round(y_Pseudo, roundZn) & "],")
                                        Next i
                                        'записываем последнюю координату
                                        writer.WriteLine("              [" & Math.Round(startPoint, roundZn) & ", " & Math.Round(endPoint, roundZn) & "]")
                                        '==========================================================================================================
                                        'записываем островки
                                        If arrayOstrov.Length > 0 Then
                                            writer.WriteLine("            ],")
                                            For i As Integer = 0 To arrayOstrov.GetUpperBound(1)
                                                Dim hg As String = arrayOstrov(1, i)
                                                Dim idOstrov As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                                If IsNothing(idOstrov) = False Then
                                                    If idOstrov.IsNull = False Then
                                                        ent = Tr.GetObject(idOstrov, OpenMode.ForWrite, False) 'открываем объект на чтение
                                                        If TypeOf ent Is Polyline Then
                                                            writer.WriteLine("            [")
                                                            Dim acPolyOstrov As Polyline = ent
                                                            Dim boolVertRedOstrov As Boolean = AcadFunction.FuncPlineDelDubleVertex(acPolyOstrov, 0.05)
                                                            Dim startPointOstrov As Double = 0
                                                            Dim endPointOstrov As Double = 0
                                                            For j As Integer = 0 To acPolyOstrov.NumberOfVertices - 1
                                                                Dim pt1 As Point2d = acPolyOstrov.GetPoint2dAt(j)
                                                                If pt1.GetVectorTo(ptPrev).Length < 0.05 Then
                                                                    Continue For
                                                                Else
                                                                    ptPrev = pt1
                                                                End If
                                                                Dim x As Double = pt1.X
                                                                Dim y As Double = pt1.Y
                                                                Dim x_Pseudo As Double = x
                                                                Dim y_Pseudo As Double = y
                                                                If IsNothing(coordSystem) = False Then
                                                                    Dim UserCoord_to_PseudoMarcator As MgCoordinateSystemTransform = coordSysFactory.GetTransform(DrawingSysCoord, coordSystem)
                                                                    Dim UserCoord_to_PseudoMercatorRez As MgCoordinate = UserCoord_to_PseudoMarcator.Transform(x, y)
                                                                    x_Pseudo = UserCoord_to_PseudoMercatorRez.X
                                                                    y_Pseudo = UserCoord_to_PseudoMercatorRez.Y
                                                                End If
                                                                If j = 0 Then
                                                                    startPointOstrov = x_Pseudo
                                                                    endPointOstrov = y_Pseudo
                                                                End If
                                                                writer.WriteLine("              [" & Math.Round(x_Pseudo, roundZn) & ", " & Math.Round(y_Pseudo, roundZn) & "],")
                                                            Next j
                                                            writer.WriteLine("              [" & Math.Round(startPointOstrov, roundZn) & ", " & Math.Round(endPointOstrov, roundZn) & " ]")
                                                            If i = arrayOstrov.GetUpperBound(1) Then
                                                                writer.WriteLine("            ]")
                                                            Else
                                                                writer.WriteLine("            ],")
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next i
                                        Else
                                            writer.WriteLine("            ]")
                                        End If
                                        If ind = ObjIDCollections.Count - 1 Then
                                            writer.WriteLine("          ]")
                                        Else
                                            writer.WriteLine("          ],")
                                        End If
                                    End If
                                Next
                                writer.WriteLine("        ]")
                                writer.WriteLine("      }")
                                writer.WriteLine("    }")
                                writer.WriteLine("  ]")
                                writer.WriteLine("}")
                            End If
                        End Using
                    Else
                        MsgBox("Сиcтема координат выбрана не верно.")
                    End If
                Catch ex As OSGeo.MapGuide.MgCoordinateSystemLoadFailedException
                    MsgBox("Не удалось найти целевую СК (WGS84.PseudoMercator)!!! Макро прервано!!!")
                    Exit Sub
                End Try
            End Using
            myStream.Close()
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'загрузка объектов в geojson для росреестра (
    <CommandMethod("ExportObjectToGeoGSON2")>
    Public Sub ExportOBJtoGeoGSON2()
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

        Dim NameDwg As String = Path.GetFileNameWithoutExtension(strDWGName)
        'временный каталог для автосохранения таблицы
        Dim TempDir As String = ""
        Try
            TempDir = System.IO.Path.GetTempPath()
        Catch ex As System.Security.SecurityException
            TempDir = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите временный каталог для автосохранения")
        End Try

        Dim rng As Random = New Random
        Dim Numbrng As Integer = rng.Next
        Dim DestDoc As String = TempDir & "Export_pkk_rosreestr_" & Numbrng & ".geojson"

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '1.считывае координатную зону чертежа
            Dim CoordSystemCode As String = CivilApplication.ActiveDocument.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
            Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
            Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
            Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
            Dim datumDict As MgCoordinateSystemDatumDictionary = csCatalog.GetDatumDictionary()
            Dim ellipsoidDict As MgCoordinateSystemEllipsoidDictionary = csCatalog.GetEllipsoidDictionary
            Dim DrawingSysCoord As MgCoordinateSystem = csDict.GetCoordinateSystem(CoordSystemCode)
            Try
                DrawingSysCoord = csDict.GetCoordinateSystem(CoordSystemCode)
            Catch ex As OSGeo.MapGuide.MgCoordinateSystemLoadFailedException
                MsgBox("Чертежу не назначена система координат!!! Макро прервано!!!")
                Exit Sub
            End Try
            '2. Пытаемся найти целевую СК
            Dim PseudoMercator As MgCoordinateSystem = csDict.GetCoordinateSystem("WGS84.PseudoMercator")
            Try

            Catch ex As OSGeo.MapGuide.MgCoordinateSystemLoadFailedException
                MsgBox("Не удалось найти целевую СК (WGS84.PseudoMercator)!!! Макро прервано!!!")
                Exit Sub
            End Try
            '3. Открываем текстовый файл на запись

            Using writer As StreamWriter = New StreamWriter(DestDoc, False)
                writer.WriteLine("{")
                writer.WriteLine(Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "FeatureCollection" & Chr(34) & ",")
                writer.WriteLine(Chr(34) & "name" & Chr(34) & ": " & Chr(34) & NameDwg & Chr(34) & ",")
                writer.WriteLine(Chr(34) & "crs" & Chr(34) & ": { " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "name" & Chr(34) & ", " & Chr(34) & "properties" & Chr(34) & ": {" & Chr(34) & "name" & Chr(34) & ": " & Chr(34) & "urn:ogc:def:crs:EPSG::3857" & Chr(34) & " } },")
                writer.WriteLine(Chr(34) & "features" & Chr(34) & ": [")
                '3. Выбираем контуры
                Dim ObjIDCollections As ObjectIdCollection = New ObjectIdCollection
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                If acSSPrompt.Status = PromptStatus.OK Then
                    Dim acSSet As SelectionSet = acSSPrompt.Value
                    For Each acSSObj As SelectedObject In acSSet
                        If Not IsDBNull(acSSObj) Then
                            ObjIDCollections.Add(acSSObj.ObjectId)
                        End If
                    Next
                End If

                Dim count As Integer = 0
                For Each id As ObjectId In ObjIDCollections
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    Dim StrCoordinates As String = ""
                    Dim StrTypeGeometry As String = ""
                    Dim StrSubClasses As String = ""
                    Dim StrLayer As String = Chr(34) & "Layer" & Chr(34) & ": " & Chr(34) & ent.Layer & Chr(34)
                    Dim StrLineType As String = Chr(34) & "Linetype" & Chr(34) & ": " & Chr(34) & ent.Linetype & Chr(34)
                    Dim StrHandle As String = Chr(34) & "EntityHandle" & Chr(34) & ": " & Chr(34) & ent.Handle.ToString & Chr(34)
                    Dim Pt3dColl As Point3dCollection = New Point3dCollection
                    Dim TypeObj As String = ""
                    Dim NameObject As String = ""
                    If TypeOf ent Is DBPoint Then
                        Dim acPoint As DBPoint = ent
                        Pt3dColl.Add(acPoint.Position)
                        StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Point" & Chr(34) & ", "
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPoint" & Chr(34)
                        TypeObj = "Point"
                    ElseIf TypeOf ent Is BlockReference Then
                        Dim acBlk As BlockReference = ent
                        Pt3dColl.Add(acBlk.Position)
                        StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Point" & Chr(34) & ", "
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPoint" & Chr(34)
                        TypeObj = "Point"
                        Dim NameStr As String = AcadFunction.FuncReadValueAttr(acBlk, "NAME")
                        If NameStr.Trim.Length > 0 Then
                            NameObject = NameStr
                        Else
                            NameObject = AcadFunction.FuncReadNameBlk(acBlk)
                        End If
                    ElseIf TypeOf ent Is CogoPoint Then
                        Dim acPointCoGo As CogoPoint = ent
                        Pt3dColl.Add(acPointCoGo.Location)
                        StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Point" & Chr(34) & ", "
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPoint" & Chr(34)
                        Dim StrName As String = acPointCoGo.PointName
                        If IsNothing(StrName) = False Then
                            If StrName.Trim.Length > 0 Then
                                NameObject = StrName
                            Else
                                StrName = acPointCoGo.FullDescription
                                If IsNothing(StrName) = False Then
                                    If StrName.Trim.Length > 0 Then
                                        NameObject = StrName
                                    End If
                                End If
                            End If
                        End If
                        TypeObj = "Point"
                    ElseIf TypeOf ent Is Circle Then
                        Dim acCircle As Circle = ent
                        Pt3dColl.Add(acCircle.Center)
                        StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Point" & Chr(34) & ", "
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPoint" & Chr(34)
                        TypeObj = "Point"
                    ElseIf TypeOf ent Is Polyline Then
                        Dim acPoly As Polyline = ent
                        For i As Integer = 0 To acPoly.NumberOfVertices - 1
                            Pt3dColl.Add(acPoly.GetPoint3dAt(i))
                        Next i
                        If acPoly.Closed = True Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        ElseIf acPoly.StartPoint.GetVectorTo(acPoly.EndPoint).Length = 0 Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        Else
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "LineString" & Chr(34) & ", "
                            TypeObj = "PLine"
                        End If
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPolyline" & Chr(34)
                    ElseIf TypeOf ent Is Polyline2d Then
                        Dim acPoly As Polyline2d = ent
                        For Each VertID As ObjectId In acPoly
                            Dim v2d As Vertex2d = acTrans.GetObject(VertID, OpenMode.ForRead)
                            Dim Point As Point3d = New Point3d(v2d.Position.X, v2d.Position.Y, 0)
                            Pt3dColl.Add(Point)
                        Next
                        If acPoly.Closed = True Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        ElseIf acPoly.StartPoint.GetVectorTo(acPoly.EndPoint).Length = 0 Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        Else
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "LineString" & Chr(34) & ", "
                            TypeObj = "PLine"
                        End If
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPolyline" & Chr(34)
                    ElseIf TypeOf ent Is Polyline3d Then
                        Dim acPoly As Polyline3d = ent
                        For Each VertID As ObjectId In acPoly
                            Dim v3d As PolylineVertex3d = acTrans.GetObject(VertID, OpenMode.ForWrite)
                            Dim Point As Point3d = New Point3d(v3d.Position.X, v3d.Position.Y, 0)
                            Pt3dColl.Add(Point)
                        Next
                        If acPoly.Closed = True Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        ElseIf acPoly.StartPoint.GetVectorTo(acPoly.EndPoint).Length = 0 Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        Else
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "LineString" & Chr(34) & ", "
                            TypeObj = "PLine"
                        End If
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPolyline" & Chr(34)
                    ElseIf TypeOf ent Is FeatureLine Then
                        Dim acPoly As FeatureLine = ent
                        Pt3dColl = acPoly.GetPoints(Autodesk.Civil.FeatureLinePointType.PIPoint)
                        If acPoly.Closed = True Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        ElseIf acPoly.StartPoint.GetVectorTo(acPoly.EndPoint).Length = 0 Then
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Polygon" & Chr(34) & ", "
                            TypeObj = "Polygon"
                        Else
                            StrTypeGeometry = Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "LineString" & Chr(34) & ", "
                            TypeObj = "PLine"
                        End If
                        StrSubClasses = Chr(34) & "SubClasses" & Chr(34) & ": " & Chr(34) & "AcDbEntity:AcDbPolyline" & Chr(34)
                    ElseIf TypeOf ent Is Parcel Then
                        Dim acPoly As Parcel = ent
                        Dim cr As Curve = acPoly.BaseCurve
                    End If
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'формируем имя объекта
                    If NameObject.Trim.Length = 0 Then
                        Dim TempStr As String = ""
                        Dim ArrayData As String(,) = Nothing
                        Dim TempName As Boolean = FuncPropertySet.FuncReadDataSetProperties(ent, "ConditionalNumber", ArrayData)
                        If IsArray(ArrayData) = False Then
                            TempName = FuncPropertySet.FuncReadDataSetProperties(ent, "CadastralNumber", ArrayData)
                            If IsArray(ArrayData) = False Then
                                TempName = FuncPropertySet.FuncReadDataSetProperties(ent, "NAME", ArrayData)
                                If IsArray(ArrayData) = False Then
                                    TempName = FuncPropertySet.FuncReadDataSetProperties(ent, "LABEL", ArrayData)
                                End If
                            End If
                        End If
                        If IsArray(ArrayData) = True Then
                            For i As Integer = 0 To ArrayData.GetUpperBound(1)
                                TempStr = ArrayData(2, i)
                                If TempStr.Trim.Length > 0 Then
                                    Exit For
                                End If
                            Next
                        End If
                        If TempStr.Trim.Length = 0 Then
                            NameObject = ent.Layer
                        Else
                            NameObject = TempStr
                        End If
                    End If
                    Dim StrNameObj As String = Chr(34) & "name" & Chr(34) & ": " & Chr(34) & NameObject & Chr(34) & ", " & Chr(34) & "buffer" & Chr(34) & ":0, "
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'читаем координаты
                    If Pt3dColl.Count > 0 Then
                        If TypeObj Like "Polygon" Then
                            Dim dl As Double = Pt3dColl.Item(0).GetVectorTo(Pt3dColl.Item(Pt3dColl.Count - 1)).Length
                            If dl <> 0 Then
                                Dim stPt As Point3d = Pt3dColl.Item(0)
                                Pt3dColl.Add(stPt)
                            End If
                        End If
                        Dim StrCoord As String = ""
                        For i As Integer = 0 To Pt3dColl.Count - 1
                            Dim x As Double = Pt3dColl.Item(i).X
                            Dim y As Double = Pt3dColl.Item(i).Y
                            Dim UserCoord_to_PseudoMarcator As MgCoordinateSystemTransform = coordSysFactory.GetTransform(DrawingSysCoord, PseudoMercator)
                            Dim UserCoord_to_PseudoMercatorRez As MgCoordinate = UserCoord_to_PseudoMarcator.Transform(x, y)
                            Dim x_Pseudo As Double = UserCoord_to_PseudoMercatorRez.X
                            Dim y_Pseudo As Double = UserCoord_to_PseudoMercatorRez.Y
                            If i = Pt3dColl.Count - 1 Then
                                StrCoord = StrCoord & "[ " & Math.Round(x_Pseudo, 3) & ", " & Math.Round(y_Pseudo, 3) & " ] "
                            Else
                                StrCoord = StrCoord & "[ " & Math.Round(x_Pseudo, 3) & ", " & Math.Round(y_Pseudo, 3) & " ], "
                            End If
                        Next i
                        Dim razd As String = ""
                        If count < ObjIDCollections.Count - 1 Then
                            razd = ","
                        End If
                        If TypeObj Like "Polygon" Then
                            writer.WriteLine("{ " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Feature" & Chr(34) & ", " & Chr(34) & "properties" & Chr(34) & ": { " & StrNameObj & StrLayer & ", " & StrSubClasses & ", " & StrHandle & " }, " & Chr(34) & "geometry" & Chr(34) & ": { " & StrTypeGeometry & Chr(34) & "coordinates" & Chr(34) & ":[[ " & StrCoord & " ]]}}" & razd)
                        ElseIf TypeObj Like "PLine" Then
                            writer.WriteLine("{ " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Feature" & Chr(34) & ", " & Chr(34) & "properties" & Chr(34) & ": { " & StrNameObj & StrLayer & ", " & StrSubClasses & ", " & StrHandle & " }, " & Chr(34) & "geometry" & Chr(34) & ": { " & StrTypeGeometry & Chr(34) & "coordinates" & Chr(34) & ":[ " & StrCoord & " ]}}" & razd)
                        Else
                            writer.WriteLine("{ " & Chr(34) & "type" & Chr(34) & ": " & Chr(34) & "Feature" & Chr(34) & ", " & Chr(34) & "properties" & Chr(34) & ": { " & StrNameObj & StrLayer & ", " & StrSubClasses & ", " & StrHandle & " }, " & Chr(34) & "geometry" & Chr(34) & ": { " & StrTypeGeometry & Chr(34) & "coordinates" & Chr(34) & ": " & StrCoord & " }}" & razd)
                        End If

                        count += 1
                    End If
                Next
                writer.WriteLine("]")
                writer.WriteLine("}")
            End Using
        End Using
        System.Diagnostics.Process.Start("explorer", TempDir)
        MsgBox("Файл формата json успешно сформирован." & vbLf & DestDoc)
    End Sub
End Class
