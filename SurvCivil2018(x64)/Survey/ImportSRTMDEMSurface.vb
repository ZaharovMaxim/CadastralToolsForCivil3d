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

Public Class ImportSRTMDEMSurface
    '===============================================================================================
    <CommandMethod("ImportSRTMDEMSurface")>
    Public Sub ImportSRTMDEMSurface()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim ed As Editor = acDoc.Editor
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
        Dim Pt3dVertexDEMColl As Point3dCollection = New Point3dCollection 'коллекция для координат вершин
        Dim NameSurfDem As String = "SurfaceDEM"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'определяем систему координат
        Dim CoordSystemCode As String = CivilApplication.ActiveDocument.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
        Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
        Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
        Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
        Dim datumDict As MgCoordinateSystemDatumDictionary = csCatalog.GetDatumDictionary()
        Dim ellipsoidDict As MgCoordinateSystemEllipsoidDictionary = csCatalog.GetEllipsoidDictionary
        Dim LL84Cs As MgCoordinateSystem = csDict.GetCoordinateSystem("LL84")
        Dim DrawingSysCoord As MgCoordinateSystem = Nothing
        Try
            DrawingSysCoord = csDict.GetCoordinateSystem(CoordSystemCode)
        Catch ex As OSGeo.MapGuide.MgCoordinateSystemLoadFailedException
            MsgBox("Чертежу не назначена система координат!!! Макро прервано!!!")
            Exit Sub
        End Try
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем файлы tiff или hgt
        '1. определяем папку с файлами dem поверхностей
        Dim DirUser As String = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку с файлами DEM", 1)
        Dim fileEntries As String() = Nothing
        If DirUser.Trim.Length > 0 Then
            Try
                fileEntries = Directory.GetFiles(DirUser)
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If

        Dim PlineBound As Polyline = Nothing
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim opt As PromptSelectionOptions = New PromptSelectionOptions()
            opt.MessageForAdding = vbLf & "Выберите проектные участки: "
            'выбираем участки
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(opt)
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            Dim tempPline As Polyline = acEnt
                            If tempPline.NumberOfVertices > 3 Then
                                PlineBound = tempPline
                            End If
                        End If
                    End If
                Next
            End If
        End Using
        '2.Читаем файлы
        If IsArray(fileEntries) = True Then
            Using acLckDoc As DocumentLock = acDoc.LockDocument()
                '5.читаем файлы tiff или hgt
                Dim count As Integer = 0
                For Each FullNameFile As String In fileEntries
                    'читаем файлы tif
                    If Path.GetExtension(FullNameFile) Like ".tif" Then
                        Dim DestCurDb As Database = Nothing
                        If count = 0 Then
                            Dim strTemplatePath As String = "acad.dwt"
                            Dim acDocMgr As DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
                            Dim DestDoc As Document = DocumentCollectionExtension.Add(acDocMgr, strTemplatePath)
                            Using acDestLok As DocumentLock = DestDoc.LockDocument()
                                DestCurDb = DestDoc.Database()
                                DestCurDb.Dimadec = 8
                                Using DestTrans As Transaction = DestCurDb.TransactionManager.StartTransaction()
                                    Dim DestCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(DestCurDb)
                                    Dim acBlkTbl As BlockTable = DestTrans.GetObject(DestCurDb.BlockTableId, OpenMode.ForWrite)
                                    Dim acBlkTblRec As BlockTableRecord = DestTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                    '3.назначаем систему координат временному чертежу WGS84
                                    DestCivilDoc.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode = "LL84"
                                    '4.назначаем временный стиль поверхности
                                    Dim userStyleID As ObjectId = DestCivilDoc.Styles.SurfaceStyles.Add("UserDEMStyle")
                                    Dim userStyle As Styles.SurfaceStyle = userStyleID.GetObject(OpenMode.ForWrite)
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Boundary).Visible = True
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Directions).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Elevations).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Gridded).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.MajorContour).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.MinorContour).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Points).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.SlopeArrows).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Slopes).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Triangles).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.UserContours).Visible = False
                                    userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Watersheds).Visible = False
                                    Try
                                        Dim gridSurfaceId As ObjectId = GridSurface.CreateFromDEM(DestCurDb, FullNameFile)
                                        Dim gridSurf As GridSurface = gridSurfaceId.GetObject(OpenMode.ForWrite)
                                        '====================================================================================
                                        'забираем все вершины в коллекцию
                                        Dim GridCelCillection As GridSurfaceVertexCollection = gridSurf.GetVertices(False)
                                        If GridCelCillection.Count > 0 Then
                                            Dim LL84WGS_to_DrawCoord As MgCoordinateSystemTransform = coordSysFactory.GetTransform(LL84Cs, DrawingSysCoord)
                                            For Each vert As GridSurfaceVertex In GridCelCillection
                                                Dim v3d As Point3d = vert.Location
                                                Dim LL84WGS_to_DrawCoordRez As MgCoordinate = LL84WGS_to_DrawCoord.Transform(v3d.X, v3d.Y)
                                                Dim x As Double = LL84WGS_to_DrawCoordRez.X
                                                Dim y As Double = LL84WGS_to_DrawCoordRez.Y
                                                If IsNothing(PlineBound) = False Then
                                                    Dim boolEntersect As Boolean = MathFunction.FuncCrossPointPolyline(PlineBound, New Point2d(x, y))
                                                    If boolEntersect = False Then
                                                        Continue For
                                                    End If
                                                End If
                                                Pt3dVertexDEMColl.Add(New Point3d(LL84WGS_to_DrawCoordRez.X, LL84WGS_to_DrawCoordRez.Y, v3d.Z))
                                            Next
                                        End If
                                    Catch ex As ArgumentException
                                    Catch ex As Autodesk.Civil.SurfaceException
                                    End Try
                                    count += 1
                                End Using
                            End Using
                        End If
                        'читаем файлы hgt
                    ElseIf Path.GetExtension(FullNameFile) Like ".hgt" Then
                        'читаем имя файла
                        Dim srtmData As Alpinechough.Srtm.ISrtmData = New Alpinechough.Srtm.SrtmData(DirUser)
                        Dim NameFile As String = Path.GetFileNameWithoutExtension(FullNameFile)
                        Dim pos1 As Integer = NameFile.IndexOf("N")
                        If pos1 = -1 Then
                            pos1 = NameFile.IndexOf("n")
                        End If
                        Dim pos2 As Integer = NameFile.IndexOf("E")
                        If pos2 = -1 Then
                            pos2 = NameFile.IndexOf("e")
                        End If
                        Dim StartLat As Double = Mid(NameFile, pos1 + 2, 2)
                        Dim StartLon As Double = Mid(NameFile, pos2 + 2, 3)
                        Dim LL84WGS_to_DrawCoord As MgCoordinateSystemTransform = coordSysFactory.GetTransform(LL84Cs, DrawingSysCoord)
                        Dim sec1 As Double = 1 / 3600
                        For i As Double = StartLat To StartLat + 1 Step sec1
                            Dim lat As Double = i
                            For j As Double = StartLon To StartLon + 1 Step sec1
                                Dim lon As Double = j
                                Dim elw As Integer = srtmData.GetElevation(New Alpinechough.Srtm.GeographicalCoordinates(lat, lon))
                                Dim LL84WGS_to_DrawCoordRez As MgCoordinate = LL84WGS_to_DrawCoord.Transform(lon, lat)
                                Dim x As Double = LL84WGS_to_DrawCoordRez.X
                                Dim y As Double = LL84WGS_to_DrawCoordRez.Y
                                If IsNothing(PlineBound) = False Then
                                    Dim boolEntersect As Boolean = MathFunction.FuncCrossPointPolyline(PlineBound, New Point2d(x, y))
                                    If boolEntersect = False Then
                                        Continue For
                                    End If
                                End If
                                Pt3dVertexDEMColl.Add(New Point3d(LL84WGS_to_DrawCoordRez.X, LL84WGS_to_DrawCoordRez.Y, elw))
                            Next j
                        Next i
                    End If
                Next
                'вспомогательная база закрыта, копируем вершины в активный чертеж
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем полилиню для запроса координат
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acBlkTbl0 As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                    Dim acBlkTblRec0 As BlockTableRecord = acTrans.GetObject(acBlkTbl0(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim StyleColl As Styles.SurfaceStyleCollection = CivilDoc.Styles.SurfaceStyles
                    If StyleColl.Count > 0 Then
                        Dim idStyleSurface As ObjectId = Nothing
                        For Each idSurfStyle As ObjectId In StyleColl
                            Dim SurfStyle As Styles.SurfaceStyle = idSurfStyle.GetObject(OpenMode.ForRead)
                            Dim Pos As Integer = SurfStyle.Name.IndexOf("Border Only")
                            If Pos > -1 Then
                                idStyleSurface = idSurfStyle
                                Exit For
                            End If
                        Next
                        If IsNothing(idStyleSurface) = True Then
                            'создаем стиль поверхности
                            idStyleSurface = CivilDoc.Styles.SurfaceStyles.Add("Border Only")
                            Dim userStyle As Styles.SurfaceStyle = idStyleSurface.GetObject(OpenMode.ForWrite)
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Boundary).Visible = True
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Directions).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Elevations).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Gridded).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.MajorContour).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.MinorContour).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Points).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.SlopeArrows).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Slopes).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Triangles).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.UserContours).Visible = False
                            userStyle.GetDisplayStylePlan(Styles.SurfaceDisplayStyleType.Watersheds).Visible = False
                        End If
                        If Pt3dVertexDEMColl.Count > 2 Then
                            Dim SurfaceDem As TinSurface = Nothing
                            Dim colSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
                            Dim TinID As ObjectId = TinSurface.Create(acCurDb, NameSurfDem)
                            Dim Tin As TinSurface = TinID.GetObject(OpenMode.ForWrite)
                            Tin.AddVertices(Pt3dVertexDEMColl)
                            If IsNothing(idStyleSurface) = False Then
                                If idStyleSurface.IsNull = False Then
                                    Tin.StyleId = idStyleSurface
                                End If
                            End If
                        End If
                    End If
                    acTrans.Commit()
                End Using
            End Using
        End If
    End Sub
End Class

