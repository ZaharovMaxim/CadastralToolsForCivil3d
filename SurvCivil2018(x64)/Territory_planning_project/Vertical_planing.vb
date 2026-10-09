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
Public Class Vertical_planing
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'точка с проектной и фактическими отметками
    <CommandMethod("VerticalPlanProjectPoint")>
    Public Sub VerticalPlanProjectPoint()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()

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
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'получаем доступ к шаблону
        Dim PutchTemplate As String = ""
        Dim NameFileTemplate As String = ""
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName 'отделяем имя файла от папки
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName 'определяем имя родительской папки для папки где находится сборка

        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG
        Dim arrayTemplates As String() = {}
        Dim countArrayPTemplates As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In IO.Directory.GetFiles(PutchTemplate)
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
        Dim formVertPoint As FormVerticalPlaningPoint = New FormVerticalPlaningPoint
        formVertPoint.NumericUpDown5.Value = MAnno
        formVertPoint.ComboBox1.DataSource = arrayTemplates
        formVertPoint.ShowDialog()
        If formVertPoint.boolShow = False Then
            Exit Sub
        End If
        MAnno = formVertPoint.NumericUpDown5.Value

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplate = parentPatch2 & "\FileResources\Sample\"
            NameFileTemplate = PutchTemplate & formVertPoint.ComboBox1.Text & ".dwg"
            If File.Exists(NameFileTemplate) = False Then
                MsgBox("Шаблон не найден. Макро прервано!!!")
                Exit Sub
            End If
        End If

        'загружаем блок
        Dim ArrayPropertiesBLK As String(,) = Nothing
        Dim NameBlk As String = ""
        Dim LayerName As String = "0"
        Dim ColorBlk As Integer = 10
        Dim ScaleBlk As Double = 1
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite)
                Dim boolReadStyleDim As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(NameFileTemplate, "Существующая и директивная отметка поверхности", ArrayPropertiesBLK, "Блок", "Вертикальная планировка территории")
                If IsArray(ArrayPropertiesBLK) = True Then
                    For i As Integer = 0 To ArrayPropertiesBLK.GetUpperBound(1)
                        If ArrayPropertiesBLK(0, i) Like "Name" Then
                            NameBlk = ArrayPropertiesBLK(1, i)
                            If IsNothing(NameBlk) = False Then
                                Dim boolCopyBLK As Boolean = AcadFunction.FuncCopyBlockDestDraw(NameFileTemplate, NameBlk)
                            End If
                        ElseIf ArrayPropertiesBLK(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayPropertiesBLK(1, i)
                            If IsNothing(NameLayerTemp) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(NameFileTemplate, NameLayerTemp)
                                If acLyrTbl.Has(NameLayerTemp) = True Then
                                    LayerName = NameLayerTemp
                                End If
                            End If
                        ElseIf ArrayPropertiesBLK(0, i) Like "Color" Then
                            ColorBlk = Val(ArrayPropertiesBLK(1, i))
                            If ColorBlk < 0 Or ColorBlk > 255 Then
                                ColorBlk = 10
                            End If
                        ElseIf ArrayPropertiesBLK(0, i) Like "Scale" Then
                            Dim OriginScale As Double = Val(ArrayPropertiesBLK(1, i))
                            If OriginScale > 0 Then
                                ScaleBlk = OriginScale * (MAnno / 1000)
                            End If
                        End If
                    Next i
                End If
                tr.Commit()
            End Using
        End Using

        'Dim CountArrayProjectSurface As Integer = 0
        Dim ProjectSurface As TinSurface = Nothing
        'Dim CountArraySurface As Integer = 0
        Dim SurfaceTIN As TinSurface = Nothing
        'Dim CountAlign As Integer = 0
        Dim Align As Alignment = Nothing

        Dim leftOffset As Double = 0
        Dim rightOffset As Double = 0
        If formVertPoint.CheckBox1.Checked = True Then
            leftOffset = -1 * formVertPoint.NumericUpDown1.Value
        End If
        If formVertPoint.CheckBox2.Checked = True Then
            rightOffset = formVertPoint.NumericUpDown2.Value
        End If

        Dim roundZn As Integer = formVertPoint.NumericUpDown3.Value
        Dim interval As Double = formVertPoint.NumericUpDown4.Value

        Using locker
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim acSSProjectTin As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Проектную поверхность: ")
                If Not (acSSProjectTin.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entProjectSurface As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSProjectTin.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entProjectSurface Is TinSurface) Then
                    MsgBox("Выбранный примитив не является Поверхностью TIN. Макро прервано.")
                    Exit Sub
                End If
                ProjectSurface = entProjectSurface

                Dim acSSTin As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Существующую поверхность: ")
                If Not (acSSTin.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entProjectTIN As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSTin.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entProjectTIN Is TinSurface) Then
                    MsgBox("Выбранный примитив не является Поверхностью TIN. Макро прервано.")
                    Exit Sub
                End If
                SurfaceTIN = entProjectTIN

                'вдоль трассы Civil 3d
                If formVertPoint.RadioButton1.Checked = True Then
                    Dim acSSAlign As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Трассу: ")
                    If Not (acSSAlign.Status = PromptStatus.OK) Then
                        Exit Sub
                    End If
                    Dim entAlign As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSAlign.ObjectId, OpenMode.ForRead)
                    If Not (TypeOf entAlign Is Alignment) Then
                        MsgBox("Выбранный примитив не является Трассой. Макро прервано.")
                        Exit Sub
                    End If
                    Align = entAlign

                    Dim IDPoly As ObjectId = Align.GetPolyline
                    Dim Poly As Polyline = tr.GetObject(IDPoly, OpenMode.ForRead)
                    Dim OldPt As Point3d = New Point3d(-1, -1, 1)
                    For i As Double = 0 To Poly.Length Step interval
                        Try
                            Dim Pt As Point3d = Poly.GetPointAtDist(i)
                            '=============================================================================================
                            'угол поворота блока
                            Dim RotBlk As Double = 0
                            If OldPt.X <> -1 And OldPt.Y <> -1 Then
                                RotBlk = Pt.Convert2d(New Plane).GetVectorTo(OldPt.Convert2d(New Plane)).Angle + Math.PI
                                OldPt = Pt
                            Else
                                OldPt = Pt
                            End If
                            If leftOffset <> 0 Then
                                Dim stationPoint As Double = -1
                                Dim offsetStation As Double = -1
                                Try
                                    Align.StationOffset(Pt.X, Pt.Y, stationPoint, offsetStation)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try

                                Dim north As Double = -1
                                Dim east As Double = -1
                                Try
                                    Align.PointLocation(stationPoint, leftOffset, east, north)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try

                                Pt = New Point3d(east, north, 0)
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = -1
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(Pt.X, Pt.Y)
                                    ProjectElev = ProjectSurface.FindElevationAtXY(Pt.X, Pt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(Pt, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            End If

                            If rightOffset <> 0 Then
                                Dim stationPoint As Double = -1
                                Dim offsetStation As Double = -1
                                Try
                                    Align.StationOffset(Pt.X, Pt.Y, stationPoint, offsetStation)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try

                                Dim north As Double = -1
                                Dim east As Double = -1

                                '==========================================================================================================
                                'вторую точку
                                Try
                                    Align.PointLocation(stationPoint, rightOffset, east, north)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                Pt = New Point3d(east, north, 0)
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = -1
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(Pt.X, Pt.Y)
                                    ProjectElev = ProjectSurface.FindElevationAtXY(Pt.X, Pt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(Pt, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            End If

                            If leftOffset = 0 And rightOffset = 0 Then
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = -1
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(Pt.X, Pt.Y)
                                    ProjectElev = ProjectSurface.FindElevationAtXY(Pt.X, Pt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(Pt, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            End If
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                        Catch ex As System.Exception
                            Exit For
                        End Try
                    Next i
                ElseIf formVertPoint.RadioButton3.Checked = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'вставка точек вручную
                    Do
                        Dim acSSPrompt As PromptPointResult = ed.GetPoint("^C^C " & vbLf & "Укажите точку: ")
                        Dim PicPoint As Point3d = acSSPrompt.Value
                        If acSSPrompt.Status = PromptStatus.OK Then
                            Dim Elev As Double = -1
                            Dim ProjectElev As Double = -1
                            Try
                                Elev = SurfaceTIN.FindElevationAtXY(PicPoint.X, PicPoint.Y)
                                ProjectElev = ProjectSurface.FindElevationAtXY(PicPoint.X, PicPoint.Y)
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                MsgBox("Не удалось найти высоту точки.")
                                Continue Do
                            End Try
                            'вставляем блок
                            Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(PicPoint, NameBlk, LayerName, ColorBlk, ScaleBlk,, 0)
                            'добавляем атрибуты
                            Dim ArrayAttrib As String(,) = Nothing
                            ReDim ArrayAttrib(1, 1)
                            ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                            ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                            ArrayAttrib(0, 1) = "ZERO_LEVEL"
                            ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                            Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                        Else
                            Exit Do
                        End If
                    Loop
                    'вдоль полилинии
                ElseIf formVertPoint.RadioButton2.Checked = True Then
                    Dim acSSAlign As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Полилинию: ")
                    If Not (acSSAlign.Status = PromptStatus.OK) Then
                        Exit Sub
                    End If
                    Dim entPoly As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSAlign.ObjectId, OpenMode.ForRead)
                    If Not (TypeOf entPoly Is Polyline) Then
                        MsgBox("Выбранный примитив не является Трассой. Макро прервано.")
                        Exit Sub
                    End If
                    Dim Poly As Polyline = entPoly
                    Dim OldPt As Point3d = New Point3d(-1, -1, 1)
                    For i As Double = 0 To Poly.Length Step interval
                        Try
                            Dim Pt As Point2d = Poly.GetPointAtDist(i).Convert2d(New Plane)
                            Dim nextInt As Double = i + interval
                            If nextInt > Poly.Length Then
                                nextInt = i - interval
                            End If
                            Dim ptNext As Point2d = New Point2d(-1, -1)
                            If nextInt < Poly.Length Then
                                ptNext = Poly.GetPointAtDist(nextInt).Convert2d(New Plane)
                            End If
                            '=============================================================================================
                            'угол поворота блока
                            Dim RotBlk As Double = 0
                            If ptNext.X <> -1 And ptNext.Y <> -1 Then
                                RotBlk = Pt.GetVectorTo(ptNext).Angle
                            End If

                            If leftOffset <> 0 Then
                                Dim leftPt As Point2d = MathFunction.FuncReturnPolarCoordinates(Pt, ptNext, RotBlk + Math.PI / 2, leftOffset)
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = -1
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(leftPt.X, leftPt.Y)
                                    ProjectElev = ProjectSurface.FindElevationAtXY(leftPt.X, leftPt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                Dim ptIns As Point3d = New Point3d(leftPt.X, leftPt.Y, 0)
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(ptIns, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            End If

                            If rightOffset <> 0 Then
                                Dim RightPt As Point2d = MathFunction.FuncReturnPolarCoordinates(Pt, ptNext, RotBlk + Math.PI / 2, rightOffset)
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = -1
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(RightPt.X, RightPt.Y)
                                    ProjectElev = ProjectSurface.FindElevationAtXY(RightPt.X, RightPt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                Dim ptIns As Point3d = New Point3d(RightPt.X, RightPt.Y, 0)
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(ptIns, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            End If

                            If leftOffset = 0 And rightOffset = 0 Then
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = -1
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(Pt.X, Pt.Y)
                                    ProjectElev = ProjectSurface.FindElevationAtXY(Pt.X, Pt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                'вставляем блок
                                Dim ptIns As Point3d = New Point3d(Pt.X, Pt.Y, 0)
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(ptIns, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            End If
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                        Catch ex As System.Exception
                            Exit For
                        End Try
                    Next i
                Else
                    Dim borders As ObjectIdCollection = New ObjectIdCollection
                    borders = ProjectSurface.ExtractBorder(SurfaceExtractionSettingsType.Plan)
                    For Each borderId As ObjectId In borders
                        Dim border As Polyline3d = borderId.GetObject(OpenMode.ForRead)
                        Dim OldPt As Point3d = New Point3d(-1, -1, 1)
                        For i As Double = 0 To border.Length Step interval
                            Try
                                Dim Pt As Point3d = border.GetPointAtDist(i)
                                Dim nextInt As Double = i + 0.01
                                If nextInt > border.Length Then
                                    nextInt = i - 0.01
                                End If
                                Dim ptNext As Point3d = border.GetPointAtDist(nextInt)
                                '=============================================================================================
                                'угол поворота блока
                                Dim RotBlk As Double = 0
                                RotBlk = Pt.Convert2d(New Plane).GetVectorTo(ptNext.Convert2d(New Plane)).Angle
                                Dim Elev As Double = -1
                                Dim ProjectElev As Double = Pt.Z
                                Try
                                    Elev = SurfaceTIN.FindElevationAtXY(Pt.X, Pt.Y)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                    Continue For
                                End Try
                                'вставляем блок
                                Dim ptIns As Point3d = New Point3d(Pt.X, Pt.Y, 0)
                                Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(ptIns, NameBlk, LayerName, ColorBlk, ScaleBlk,, RotBlk)
                                'добавляем атрибуты
                                Dim ArrayAttrib As String(,) = Nothing
                                ReDim ArrayAttrib(1, 1)
                                ArrayAttrib(0, 0) = "PROJECT_LEVEL"
                                ArrayAttrib(1, 0) = FuncFormatZN.FuncFormatCoordinate(ProjectElev, roundZn)
                                ArrayAttrib(0, 1) = "ZERO_LEVEL"
                                ArrayAttrib(1, 1) = FuncFormatZN.FuncFormatCoordinate(Elev, roundZn)
                                Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            Catch ex As System.Exception
                                Exit For
                            End Try
                        Next i
                    Next
                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'расстановка уклонов на основе продольного профиля
    <CommandMethod("VerticalPlanSlopeDist")>
    Public Sub VerticalPlanSlopeDist()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()

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
        Dim MAnno As Integer = AcadFunction.FuncReadCurrentAnnotativeScaleDWG

        Dim formVertPlan As FormVerticalPlaningSlopeDist = New FormVerticalPlaningSlopeDist
        formVertPlan.NumericUpDown1.Value = MAnno
        formVertPlan.ShowDialog()
        If formVertPlan.boolShov = False Then
            Exit Sub
        End If
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'получаем доступ к шаблону
        Dim PutchTemplate As String = ""
        Dim NameFileTemplate As String = ""
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName 'отделяем имя файла от папки
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName 'определяем имя родительской папки для папки где находится сборка
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'файл-шаблон
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In IO.Directory.GetFiles(PutchTemplate)
                Try
                    Dim File1 As String = IO.Path.GetFileName(fileDwg)
                    Dim ext As String = IO.Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        NameFileTemplate = File1
                        Exit For
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        'загружаем блок
        Dim ArrayPropertiesBLK As String(,) = Nothing
        Dim NameBlk As String = ""
        Dim LayerName As String = "0"
        Dim ColorBlk As Integer = 10
        Dim ScaleBlk As Double = 1

        Dim ArrayPropertiesMleader As String(,) = Nothing
        Dim IDMLeaderStyle As ObjectId = Nothing
        Dim NameStyleMLeader As String = ""
        Dim LayerNameMLeader As String = "0"
        Dim HTextMLeader As Double = 3
        Dim ColorMLeader As Integer = 7
        Dim SizeStrelka As Double = 0

        Dim offsetMLeader As Double = formVertPlan.NumericUpDown3.Value
        Dim heightText As Double = formVertPlan.NumericUpDown2.Value
        Dim M As Double = formVertPlan.NumericUpDown1.Value
        Dim kf As Integer = 1
        If formVertPlan.RadioButton3.Checked = True Then
            kf = -1
        End If
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForWrite)
                Dim mlstyles As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForWrite)
                Dim boolReadStyleDim As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate & NameFileTemplate, "Уклон-Расстояние", ArrayPropertiesBLK, "Блок", "Вертикальная планировка территории")
                If IsArray(ArrayPropertiesBLK) = True Then
                    For i As Integer = 0 To ArrayPropertiesBLK.GetUpperBound(1)
                        If ArrayPropertiesBLK(0, i) Like "Name" Then
                            Dim TempNameBlk As String = ArrayPropertiesBLK(1, i)
                            If IsNothing(TempNameBlk) = False Then
                                Dim boolCopyBLK As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchTemplate & NameFileTemplate, TempNameBlk)
                                If acBlkTbl.Has("Уклон-Расстояние") = False Then
                                    NameBlk = TempNameBlk
                                End If
                            End If
                        ElseIf ArrayPropertiesBLK(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayPropertiesBLK(1, i)
                            If IsNothing(NameLayerTemp) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplate & NameFileTemplate, NameLayerTemp)
                                If acLyrTbl.Has(NameLayerTemp) = True Then
                                    LayerName = NameLayerTemp
                                End If
                            End If
                        ElseIf ArrayPropertiesBLK(0, i) Like "Color" Then
                            ColorBlk = Val(ArrayPropertiesBLK(1, i))
                            If ColorBlk < 0 Or ColorBlk > 255 Then
                                ColorBlk = 10
                            End If
                        ElseIf ArrayPropertiesBLK(0, i) Like "Scale" Then
                            Dim OriginScale As Double = Val(ArrayPropertiesBLK(1, i))
                            If OriginScale > 0 Then
                                ScaleBlk = OriginScale * (M / 1000)
                            End If
                        End If
                    Next i
                End If

                If NameBlk.Trim.Length = 0 Then
                    MsgBox("Блок УКЛОН-РАССТОЯНИЕ не найден начертеже. МАкро прервано!!!")
                    locker.Dispose()
                    Exit Sub
                End If

                Dim boolReadStyleCurve As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate & NameFileTemplate, "Пояснительная надпись", ArrayPropertiesMleader, "Мультивыноска", "Трасса автомобильной дороги")
                If IsArray(ArrayPropertiesMleader) = True Then
                    For i As Integer = 0 To ArrayPropertiesMleader.GetUpperBound(1)
                        If ArrayPropertiesMleader(0, i) Like "Name" Then
                            Dim NameMleaderStyleCurve As String = ArrayPropertiesMleader(1, i)
                            If IsNothing(NameMleaderStyleCurve) = False Then
                                Dim boolCopyMleaderStyle As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplate & NameFileTemplate, NameMleaderStyleCurve)
                                For Each item As DBDictionaryEntry In mlstyles
                                    If NameMleaderStyleCurve Like item.Key Then
                                        IDMLeaderStyle = item.Value
                                        Exit For
                                    End If
                                Next
                            End If
                        ElseIf ArrayPropertiesMleader(0, i) Like "Layer" Then
                            Dim NameLayerTemp As String = ArrayPropertiesMleader(1, i)
                            If IsNothing(NameLayerTemp) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplate & NameFileTemplate, NameLayerTemp)
                                If acLyrTbl.Has(NameLayerTemp) = True Then
                                    LayerNameMLeader = NameLayerTemp
                                    Exit For
                                End If
                            End If
                        ElseIf ArrayPropertiesMleader(0, i) Like "Color" Then
                            ColorMLeader = Val(ArrayPropertiesMleader(1, i))
                            If ColorMLeader < 0 Or ColorMLeader > 255 Then
                                ColorMLeader = 10
                            End If
                        ElseIf ArrayPropertiesMleader(0, i) Like "TextHeight" Then
                            HTextMLeader = Val(ArrayPropertiesMleader(1, i))
                            If HTextMLeader < 0 Then
                                HTextMLeader = heightText
                            End If
                        ElseIf ArrayPropertiesMleader(0, i) Like "ArrowSize" Then
                            SizeStrelka = Val(ArrayPropertiesMleader(1, i))
                            If SizeStrelka < 0 Then
                                SizeStrelka = 0
                            End If
                        End If
                    Next i
                End If
                tr.Commit()
            End Using
            If IsNothing(IDMLeaderStyle) = True Then
                MsgBox("Мультивноска не найдена. МАкро прервано!!!")
                locker.Dispose()
                Exit Sub
            End If

            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim Align As Alignment = Nothing
            tr = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Dim acSSAlign As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Трассу Civil3D: ")
                If Not (acSSAlign.Status = PromptStatus.OK) Then
                    Exit Sub
                End If
                Dim entAlign As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSAlign.ObjectId, OpenMode.ForRead)
                If Not (TypeOf entAlign Is Alignment) Then
                    MsgBox("Выбранный примитив не является Трассой. Макро прервано.")
                    Exit Sub
                End If
                Align = entAlign

                If formVertPlan.RadioButton1.Checked = True Then
                    '==============================================================================================================
                    'начальный пикет
                    Dim optPoint1 As PromptPointOptions = New PromptPointOptions("Укажите начальную точку трассы:")
                    Dim PPrPointAlign As PromptPointResult = ed.GetPoint(optPoint1)
                    If Not (PPrPointAlign.Status = PromptStatus.OK) Then
                        locker.Dispose()
                        Exit Sub
                    End If
                    Dim StartPointAlign As Point3d = PPrPointAlign.Value
                    Dim StartPK As Double = -1
                    Dim PKOffset As Double = -1
                    Try
                        Align.StationOffset(StartPointAlign.X, StartPointAlign.Y, StartPK, PKOffset)
                    Catch ex As Exception
                        MsgBox("Не удалось определить начальный пикет на трассе!!! Макро прервано.")
                        locker.Dispose()
                        Exit Sub
                    End Try
                    'начальная точка на продольном плофиле
                    Dim StartPosPtofileX As Double = 0
                    Dim flag As Boolean = False 'флаг, запоминаем начальную точку профиля
                    Do
                        '==============================================================================================================
                        'начальный пикет
                        Dim optPoint2 As PromptPointOptions = New PromptPointOptions("Укажите начальную точку на продольном профиле:")
                        Dim PPrStartPointProfile As PromptPointResult = ed.GetPoint(optPoint2)
                        If Not (PPrStartPointProfile.Status = PromptStatus.OK) Then
                            Exit Do
                        End If
                        Dim StartPointProfile As Point3d = PPrStartPointProfile.Value
                        If flag = False Then
                            StartPosPtofileX = PPrStartPointProfile.Value.X
                            flag = True
                        End If
                        '==============================================================================================================
                        'конечный пикет
                        Dim optPoint3 As PromptPointOptions = New PromptPointOptions("Укажите конечную точку на продольном профиле:")
                        Dim PPrEndPointProfile As PromptPointResult = ed.GetPoint(optPoint3)
                        If Not (PPrEndPointProfile.Status = PromptStatus.OK) Then
                            Exit Do
                        End If
                        Dim EndPointProfile As Point3d = PPrEndPointProfile.Value
                        '==============================================================================================================
                        'вычисляем расстояние
                        Dim DistStart As Double = Math.Abs(StartPosPtofileX - StartPointProfile.X)
                        Dim DistEnd As Double = Math.Abs(StartPosPtofileX - EndPointProfile.X)
                        Dim DistSr As Double = (DistEnd + DistStart) / 2

                        '==============================================================================================================
                        'уклон
                        Dim pStrOpts2 As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите уклон: ")
                        Dim pStrRes2 As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts2)
                        If Not (pStrRes2.Status = PromptStatus.OK) Then
                            Exit Do
                        End If
                        Dim UserSlope As Double = pStrRes2.Value
                        '===============================================================================================================
                        'вычисляем пикет для вставки уклона
                        Dim PKInsert As Double = StartPK + DistSr
                        'вычисляем координаты вставки блока
                        Dim x As Double = -1
                        Dim y As Double = -1
                        Dim xboff As Double = -1
                        Dim yboff As Double = -1
                        Try
                            Align.PointLocation(PKInsert, 0, x, y)
                            Align.PointLocation(PKInsert + 0.01, 0, xboff, yboff)
                            Dim PTInsert As Point3d = New Point3d(x, y, 0)
                            Dim PTRotate As Point2d = New Point2d(xboff, yboff)
                            Dim root As Double = PTInsert.Convert2d(New Plane).GetVectorTo(PTRotate).Angle
                            'вставляем блок
                            Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(PTInsert, NameBlk, LayerName, ColorBlk, ScaleBlk, 20, root)
                            'добавляем атрибуты
                            Dim ArrayAttrib As String(,) = Nothing
                            ReDim ArrayAttrib(1, 1)
                            ArrayAttrib(0, 0) = "Уклон"
                            ArrayAttrib(1, 0) = MathFunction.FuncFormatNumb(UserSlope, 2)
                            ArrayAttrib(0, 1) = "РАССТОЯНИЕ"
                            Dim distCalk As Double = Math.Abs(EndPointProfile.X - StartPointProfile.X)
                            ArrayAttrib(1, 1) = MathFunction.FuncFormatNumb(distCalk, 2)
                            Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                        End Try
                        'расставляем мультивыноски с пикетами
                        If IsNothing(IDMLeaderStyle) = False Then
                            Dim xStart As Double = -1
                            Dim yStart As Double = -1
                            Dim xoff As Double = -1
                            Dim yoff As Double = -1
                            Try
                                'получили точку вставки мультивыноски
                                Align.PointLocation(StartPK + DistStart, 0, xStart, yStart)
                                Align.PointLocation(StartPK + DistStart, kf * offsetMLeader, xoff, yoff)
                                Dim insPt As Point2d = New Point2d(xStart, yStart)
                                Try
                                    Dim startPKstr As String = Align.GetStationStringWithEquations(StartPK + DistStart)
                                    Dim idMleader As ObjectId = AcadFunction.FuncDrawMleader(insPt, "ПК " & startPKstr, xoff - xStart, yoff - yStart, LayerNameMLeader, HTextMLeader, IDMLeaderStyle,, ColorMLeader, False, True)
                                Catch ex As System.ArgumentException
                                End Try
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            End Try
                        End If
                    Loop
                ElseIf formVertPlan.RadioButton2.Checked = True Then
                    Dim acSSProfile As PromptEntityResult = ed.GetEntity(vbLf & "Выберите линию Профиля Civil3D: ")
                    If Not (acSSProfile.Status = PromptStatus.OK) Then
                        Exit Sub
                    End If
                    Dim entProfile As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(acSSProfile.ObjectId, OpenMode.ForRead)
                    If Not (TypeOf entProfile Is Profile) Then
                        MsgBox("Выбранный примитив не является Профилем. Макро прервано.")
                        Exit Sub
                    End If
                    Dim Prof As Profile = entProfile
                    Dim profColl As ProfileEntityCollection = Prof.Entities
                    Dim arrayElement As Double(,) = {}
                    Dim countArrayElement As Integer = 0
                    If profColl.Count > 0 Then
                        For i As Integer = 0 To profColl.Count - 1
                            Dim entPr As ProfileEntity = profColl.Item(i)
                            If entPr.EntityType = ProfileEntityType.Tangent Then
                                Dim TEnt As ProfileTangent = entPr
                                ReDim Preserve arrayElement(6, countArrayElement)
                                arrayElement(0, countArrayElement) = TEnt.StartStation
                                arrayElement(1, countArrayElement) = TEnt.EndStation
                                arrayElement(2, countArrayElement) = TEnt.StartElevation
                                arrayElement(3, countArrayElement) = TEnt.EndElevation
                                arrayElement(4, countArrayElement) = TEnt.Length
                                arrayElement(5, countArrayElement) = 0
                                arrayElement(6, countArrayElement) = 0
                                countArrayElement += 1
                            ElseIf entPr.EntityType = ProfileEntityType.Circular Then
                                Dim CEnt As ProfileCircular = entPr
                                ReDim Preserve arrayElement(6, countArrayElement)
                                arrayElement(0, countArrayElement) = CEnt.StartStation
                                arrayElement(1, countArrayElement) = CEnt.EndStation
                                arrayElement(2, countArrayElement) = CEnt.StartElevation
                                arrayElement(3, countArrayElement) = CEnt.EndElevation
                                arrayElement(4, countArrayElement) = CEnt.Length
                                Try
                                    Dim stat As Double = (CEnt.EndStation + CEnt.StartStation) / 2
                                    arrayElement(5, countArrayElement) = Prof.ElevationAt(stat)
                                    arrayElement(6, countArrayElement) = stat
                                Catch ex As System.ArgumentException
                                    arrayElement(5, countArrayElement) = 0
                                    arrayElement(6, countArrayElement) = 0
                                End Try
                                countArrayElement += 1
                            ElseIf entPr.EntityType = ProfileEntityType.ParabolaAsymmetric Then
                                Dim AEnt As ProfileParabolaAsymmetric = entPr
                                ReDim Preserve arrayElement(6, countArrayElement)
                                arrayElement(0, countArrayElement) = AEnt.StartStation
                                arrayElement(1, countArrayElement) = AEnt.EndStation
                                arrayElement(2, countArrayElement) = AEnt.StartElevation
                                arrayElement(3, countArrayElement) = AEnt.EndElevation
                                arrayElement(4, countArrayElement) = AEnt.Length
                                Try
                                    Dim stat As Double = (AEnt.EndStation + AEnt.StartStation) / 2
                                    arrayElement(5, countArrayElement) = Prof.ElevationAt(stat)
                                    arrayElement(6, countArrayElement) = stat
                                Catch ex As System.ArgumentException
                                    arrayElement(5, countArrayElement) = 0
                                    arrayElement(6, countArrayElement) = 0
                                End Try
                                countArrayElement += 1
                            ElseIf entPr.EntityType = ProfileEntityType.ParabolaSymmetric Then
                                Dim AEnt As ProfileParabolaSymmetric = entPr
                                ReDim Preserve arrayElement(6, countArrayElement)
                                arrayElement(0, countArrayElement) = AEnt.StartStation
                                arrayElement(1, countArrayElement) = AEnt.EndStation
                                arrayElement(2, countArrayElement) = AEnt.StartElevation
                                arrayElement(3, countArrayElement) = AEnt.EndElevation
                                arrayElement(4, countArrayElement) = AEnt.Length
                                Try
                                    Dim stat As Double = (AEnt.EndStation + AEnt.StartStation) / 2
                                    arrayElement(5, countArrayElement) = Prof.ElevationAt(stat)
                                    arrayElement(6, countArrayElement) = stat
                                Catch ex As System.ArgumentException
                                    arrayElement(5, countArrayElement) = 0
                                    arrayElement(6, countArrayElement) = 0
                                End Try
                                countArrayElement += 1
                            End If
                        Next
                    End If
                    If arrayElement.Length > 0 Then
                        Dim boolSortArray As Boolean = MathFunction.FuncSortDblArray(arrayElement, 0)
                    End If
                    For i As Integer = 0 To arrayElement.GetUpperBound(1)
                        Dim k As Double = arrayElement(5, i)
                        Dim iSlope As Double = 0
                        If k = 0 Then
                            iSlope = (arrayElement(3, i) - arrayElement(2, i)) / arrayElement(4, i)
                            iSlope = Math.Abs(iSlope * 1000)
                        Else
                            Dim tempI1 As Double = (arrayElement(2, i) - arrayElement(5, i)) / Math.Abs((arrayElement(6, i) - arrayElement(0, i)))
                            Dim tempI2 As Double = (arrayElement(3, i) - arrayElement(5, i)) / Math.Abs((arrayElement(6, i) - arrayElement(1, i)))
                            If tempI1 > tempI2 Then
                                iSlope = Math.Abs(tempI1 * 1000)
                            Else
                                iSlope = Math.Abs(tempI2 * 1000)
                            End If
                        End If
                        '===============================================================================================================
                        'вычисляем пикет для вставки уклона
                        Dim PKInsert As Double = (arrayElement(0, i) + arrayElement(1, i)) / 2
                        'вычисляем координаты вставки блока
                        Dim x As Double = -1
                        Dim y As Double = -1
                        Dim xboff As Double = -1
                        Dim yboff As Double = -1
                        Try
                            Align.PointLocation(PKInsert, 0, x, y)
                            If iSlope > 0 Then
                                Align.PointLocation(PKInsert + 0.01, 0, xboff, yboff)
                            Else
                                Align.PointLocation(PKInsert - 0.01, 0, xboff, yboff)
                            End If
                            Dim PTInsert As Point3d = New Point3d(x, y, 0)
                            Dim PTOrient As Point2d = New Point2d(xboff, yboff)
                            Dim rot As Double = PTInsert.Convert2d(New Plane).GetVectorTo(PTOrient).Angle
                            'вставляем блок
                            Dim Blk As BlockReference = AcadFunction.FuncInsertDrawingBlk(PTInsert, NameBlk, LayerName, ColorBlk, ScaleBlk, 20, rot, 1)
                            'добавляем атрибуты
                            Dim ArrayAttrib As String(,) = Nothing
                            ReDim ArrayAttrib(1, 1)
                            ArrayAttrib(0, 0) = "Уклон"
                            ArrayAttrib(1, 0) = MathFunction.FuncFormatNumb(iSlope, 2)
                            ArrayAttrib(0, 1) = "РАССТОЯНИЕ"
                            ArrayAttrib(1, 1) = MathFunction.FuncFormatNumb(arrayElement(4, i), 2)
                            Dim boolAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(Blk, ArrayAttrib)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                        End Try

                        'расставляем мультивыноски с пикетами
                        If IsNothing(IDMLeaderStyle) = False Then
                            Dim xStart As Double = -1
                            Dim yStart As Double = -1
                            Dim xoff As Double = -1
                            Dim yoff As Double = -1
                            Try
                                'получили точку вставки мультивыноски
                                Align.PointLocation(arrayElement(0, i), 0, xStart, yStart)
                                Align.PointLocation(arrayElement(0, i), kf * offsetMLeader, xoff, yoff)
                                Dim insPt As Point2d = New Point2d(xStart, yStart)
                                Try
                                    Dim startPKstr As String = Align.GetStationStringWithEquations(arrayElement(0, i))
                                    Dim idMleader As ObjectId = AcadFunction.FuncDrawMleader(insPt, "ПК " & startPKstr, 20, 20, LayerNameMLeader, HTextMLeader, IDMLeaderStyle,, ColorMLeader, False, True)
                                    'Dim idMleader As ObjectId = AcadFunction.FuncDrawMleader(insPt, "ПК " & startPKstr, xoff - xStart, yoff - yStart, LayerNameMLeader, HTextMLeader, IDMLeaderStyle,, ColorMLeader, False, True)
                                Catch ex As System.ArgumentException
                                End Try
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                            End Try
                            If i = arrayElement.GetUpperBound(1) Then
                                Dim xEnd As Double = -1
                                Dim yEnd As Double = -1
                                Try
                                    'получили точку вставки мультивыноски
                                    Align.PointLocation(arrayElement(1, i), 0, xEnd, yEnd)
                                    Dim insPt As Point2d = New Point2d(xEnd, yEnd)
                                    Try
                                        Dim endPKstr As String = Align.GetStationStringWithEquations(arrayElement(1, i))
                                        Dim idMleader As ObjectId = AcadFunction.FuncDrawMleader(insPt, "ПК " & endPKstr, 20, 20, LayerNameMLeader, HTextMLeader, IDMLeaderStyle,, ColorMLeader, False, True)
                                        'Dim idMleader As ObjectId = AcadFunction.FuncDrawMleader(insPt, "ПК " & endPKstr, xoff - xStart, yoff - yStart, LayerNameMLeader, HTextMLeader, IDMLeaderStyle,, ColorMLeader, False, True)
                                    Catch ex As System.ArgumentException
                                    End Try
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            End If
                        End If
                    Next i
                End If
                tr.Commit()
            End Using
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

End Class
