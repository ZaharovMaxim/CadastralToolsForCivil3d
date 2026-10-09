Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Public Class HatchingContour
    '=============================================================================================
    'заполнение контура (штриховка)
    <CommandMethod("HatchingContour")>
    Public Sub HatchingContour()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        '===========================================================================
        'типы объектов (всего 3 типа для обработки, полилиния, 2д полилиния и окружность)
        Dim PObl As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing 'полилиния
        Dim P2Obl As Polyline2d = Nothing '2d полилиния
        Dim CircObj As Circle = Nothing
        '=========================================================================
        'вызываем диалоговое окно
        Dim FormHatchCounter As HatchContour = New HatchContour
        FormHatchCounter.ShowDialog()
        If FormHatchCounter.stbool = False Then Exit Sub
        '=========================================================================
        'блок для заполнения контура
        Dim PutchTemplate As String = FormHatchCounter.PutchBLK & FormHatchCounter.ComboBox3.Text
        If File.Exists(PutchTemplate) = False Then
            MsgBox("Шаблон не найден. Макро прервано!!!")
            Exit Sub
        End If
        Dim DeskBlk As String = FormHatchCounter.ComboBox2.Text
        Dim ArrayPropBlk As String(,) = Nothing
        Dim NameTable As String = "Растительность"
        If FormHatchCounter.RadioButton4.Checked = True Then
            NameTable = "Растительность"
        Else
            NameTable = "Сельскохозяйственные угодья"
        End If
        Dim BLK As BlockReference = Nothing
        Dim MVBlock As Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition = Nothing
        Dim mvb As Autodesk.Aec.DatabaseServices.MultiViewBlockReference = Nothing
        Dim NameBlk As String = ""
        Dim MScale As Double = 1
        Dim Rot As Double = 0
        Dim NameLayer As String = "0"
        Dim indColor As Integer = 7
        Dim ScaleMVB As Scale3d = New Scale3d(FormHatchCounter.NumericUpDown4.Value, FormHatchCounter.NumericUpDown4.Value, FormHatchCounter.NumericUpDown4.Value)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
            Dim mvbstyledict As Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition = New Autodesk.Aec.DatabaseServices.DictionaryMultiViewBlockDefinition(acCurDb)
            Dim userSurface As TinSurface = Nothing
            If FormHatchCounter.CheckBox1.Checked = True Then
                If CivilDoc.GetSurfaceIds.Count > 0 Then
                    Dim IDSurf As ObjectId = CivilDoc.GetSurfaceIds.Item(0)
                    If IDSurf.IsNull = False Then
                        userSurface = acTrans.GetObject(IDSurf, OpenMode.ForRead)
                    End If
                Else
                    MsgBox("В чертеже нет поверхности TIN. Макро прервано!!!")
                    Exit Sub
                End If

            End If
            If FormHatchCounter.CheckBox1.Checked = False Then
                If FormHatchCounter.boolSelectBlk = False Then
                    If File.Exists(PutchTemplate) = True Then
                        Dim boolPropBlkUPK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplate, DeskBlk, ArrayPropBlk, "Блок", NameTable)
                        If IsArray(ArrayPropBlk) = True Then
                            For k As Integer = 0 To ArrayPropBlk.GetUpperBound(1)
                                Dim NameField As String = ArrayPropBlk(0, k)
                                If NameField.Trim Like "Name" Then
                                    NameBlk = ArrayPropBlk(1, k)
                                    If IsNothing(NameBlk) = False Then
                                        If NameBlk.Trim.Length > 0 Then
                                            Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchTemplate, NameBlk)
                                            If acBlkTbl.Has(NameBlk) = False Then
                                                MsgBox("Определение блока в шаблоне не найдено. Макро прервано!!!")
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                ElseIf NameField.Trim Like "Color" Then
                                    indColor = Val(ArrayPropBlk(1, k))
                                    If indColor < 0 Then indColor = 0
                                    If indColor > 2560 Then indColor = 256
                                ElseIf NameField.Trim Like "Layer" Then
                                    NameLayer = Val(ArrayPropBlk(1, k))
                                    If acLyrTbl.Has(NameLayer) = False Then
                                        Dim boolInsLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplate, NameLayer)
                                        If acLyrTbl.Has(NameLayer) = False Then
                                            NameLayer = "0"
                                        End If
                                    End If
                                ElseIf NameField.Trim Like "Scale" Then
                                    Dim TempScaleBlk As Double = Val(ArrayPropBlk(1, k))
                                    Dim MDwg As Double = FormHatchCounter.NumericUpDown3.Value
                                    MScale = TempScaleBlk * (MDwg / 1000)
                                End If
                            Next k
                        Else
                            MsgBox("Характеристики блока не найдены. Макро прервано!!!")
                            Exit Sub
                        End If
                    End If
                Else
                    Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(FormHatchCounter.idBlk, OpenMode.ForRead)
                    If Not IsDBNull(acEnt) Then
                        If TypeOf acEnt Is BlockReference Then
                            BLK = acEnt
                            If BLK.IsDynamicBlock = True Then
                                Dim blockRef As BlockReference = acTrans.GetObject(FormHatchCounter.idBlk, OpenMode.ForRead)
                                Dim blockDyn As BlockTableRecord = acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
                                NameLayer = BLK.Layer
                                NameBlk = blockDyn.Name
                                indColor = BLK.ColorIndex
                                Rot = BLK.Rotation
                                MScale = BLK.ScaleFactors.X
                            Else
                                NameLayer = BLK.Layer
                                NameBlk = BLK.Name
                                Rot = BLK.Rotation
                                indColor = BLK.ColorIndex
                                MScale = BLK.ScaleFactors.X
                            End If
                        Else
                            MsgBox("Вабранный объект не является блоком. Мокрос завершен")
                            Exit Sub
                        End If
                    End If
                End If
            Else
                'штрихуем 3д блоками
                'считываем блок из файла и вставляем его определение
                PutchTemplate = FormHatchCounter.PutchMVBLK & "MVBlocks.dwg"
                If IO.File.Exists(PutchTemplate) = True Then
                    Dim boolcopyMVBlock As Boolean = AcadFunction.FuncCopyMVBlockDestDraw(PutchTemplate, FormHatchCounter.ComboBox2.Text.Trim, False)
                    If boolcopyMVBlock = True Then

                    End If
                End If
            End If
            '==============================================================================
            'делаем расчет сетки и определяем масштаб чертежа
            Dim Shag As Double = FormHatchCounter.NumericUpDown1.Value
            Dim m As Double = FormHatchCounter.NumericUpDown3.Value
            If m = 0 Then m = 1000
            Shag = (m / 100) * Shag
            If m <> 1000 Then
                Dim m1 As Double = m / 1000
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем объект для заполнения
            Dim acSSPrompt As PromptSelectionResult = ed.GetSelection()
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each acSSObj As SelectedObject In acSSet
                    If Not IsDBNull(acSSObj) Then
                        '=======================================================================================================================
                        'выбираем внешний контур
                        Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(obj) Then
                            If obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline" Then
                                PObl = obj
                                PObl.Closed = True
                            ElseIf obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline2d" Then
                                P2Obl = obj
                                P2Obl.Closed = True
                            ElseIf obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Circle" Then
                                CircObj = obj
                            Else
                                Continue For
                            End If
                            'читаем расширенные данные
                            Dim ArrayData As String(,) = Nothing
                            Dim boolReadData As Boolean = FuncPropertySet.FuncReadAllDataSetProperties(obj, ArrayData)
                            If IsArray(ArrayData) = True Then
                                If FormHatchCounter.NumericUpDown4.Value = 1 Then
                                    For j1 As Integer = 0 To ArrayData.GetUpperBound(1)
                                        Dim NameField As String = ArrayData(1, j1)
                                        If NameField.Trim Like "HEIGHT" Then
                                            Dim TempH As Double = Val(ArrayData(2, j1))
                                            TempH = TempH / MScale
                                            ScaleMVB = New Scale3d(TempH, TempH, TempH)
                                        End If
                                    Next j1
                                End If
                            End If
                            '=============================================================================================================
                            'находим габаритный контейнер
                            Dim MinCoord As Point3d = Nothing 'находим габаритный контейнер
                            Dim MaxCoord As Point3d = Nothing
                            If IsNothing(PObl) = False Then
                                MinCoord = PObl.GeometricExtents.MinPoint 'находим габаритный контейнер
                                MaxCoord = PObl.GeometricExtents.MaxPoint
                            ElseIf IsNothing(P2Obl) = False Then
                                MinCoord = P2Obl.GeometricExtents.MinPoint 'находим габаритный контейнер
                                MaxCoord = P2Obl.GeometricExtents.MaxPoint
                            ElseIf IsNothing(CircObj) = False Then
                                MinCoord = CircObj.GeometricExtents.MinPoint 'находим габаритный контейнер
                                MaxCoord = CircObj.GeometricExtents.MaxPoint
                            End If
                            Dim MinX As Double = Math.Round(MinCoord.X, 3)
                            Dim MinY As Double = Math.Round(MinCoord.Y, 3)
                            Dim dl As Double = MaxCoord.X - MinCoord.X 'длина прямоугольника
                            Dim sh As Double = MaxCoord.Y - MinCoord.Y 'ширина прямоугольника
                            Dim i As Integer = dl / Shag
                            Dim j As Integer = sh / Shag
                            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                            'штриховка по квадратам
                            Dim gd As DBDictionary = acTrans.GetObject(acCurDb.GroupDictionaryId, OpenMode.ForWrite)
                            Dim ids As ObjectIdCollection = New ObjectIdCollection()
                            If FormHatchCounter.RadioButton1.Checked = True Then
                                For j1 As Integer = 0 To i
                                    For i1 = 0 To j
                                        'координаты точки вставки блока
                                        Dim XPt As Double = Math.Round(MinCoord.X + j1 * Shag, 3)
                                        Dim YPt As Double = Math.Round(MinCoord.Y + i1 * Shag, 3)
                                        Dim Pt1 As Point3d = New Point3d(XPt, YPt, 0)
                                        'флаг, который указывает на то что точка лежин внутри внешнего контура
                                        'блок для проверки условия
                                        Dim insBoolPoint As Boolean = False
                                        If IsNothing(PObl) = False Then
                                            insBoolPoint = MathFunction.FuncCrossPointPolygon(PObl, Pt1)
                                        ElseIf IsNothing(P2Obl) = False Then
                                            insBoolPoint = MathFunction.FuncCrossPointPolygon(P2Obl, Pt1)
                                        ElseIf IsNothing(CircObj) = False Then
                                            insBoolPoint = MathFunction.FuncCrossPointPolygon(CircObj, Pt1)
                                        End If
                                        If insBoolPoint = True Then
                                            Dim InsPoint As Point3d = New Point3d(XPt, YPt, 0)
                                            If FormHatchCounter.CheckBox1.Checked = False Then
                                                Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPoint, NameBlk, NameLayer, indColor, MScale, 20, Rot)
                                                ids.Add(InsBlk.Id)
                                            Else
                                                Dim idmvbstyledict As ObjectId = mvbstyledict.GetAt(FormHatchCounter.ComboBox2.Text.Trim)
                                                mvb = New Autodesk.Aec.DatabaseServices.MultiViewBlockReference()

                                                mvb.BlockDefId = idmvbstyledict
                                                If IsNothing(userSurface) = False Then
                                                    Try
                                                        Dim elev As Double = userSurface.FindElevationAtXY(XPt, YPt)
                                                        InsPoint = New Point3d(XPt, YPt, elev)
                                                        mvb.Location = InsPoint
                                                        mvb.Layer = obj.Layer
                                                        mvb.Scale = ScaleMVB
                                                        mvb.SetDatabaseDefaults()
                                                        acBlkTblRec.AppendEntity(mvb)
                                                        acTrans.AddNewlyCreatedDBObject(mvb, True)
                                                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                    End Try
                                                End If
                                            End If
                                        End If
                                    Next i1
                                Next j1
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'штриховка в шахматном порядке
                            ElseIf FormHatchCounter.RadioButton2.Checked = True Then
                                For j1 As Integer = 0 To i
                                    For i1 = 0 To j
                                        If (j1 Mod 2) = 0 Then
                                            If (i1 Mod 2) <> 0 Then
                                                'координаты точки вставки блока
                                                Dim XPt As Double = Math.Round(MinCoord.X + j1 * Shag, 3)
                                                Dim YPt As Double = Math.Round(MinCoord.Y + i1 * Shag, 3)
                                                Dim Pt1 As Point3d = New Point3d(XPt, YPt, 0)
                                                'флаг, который указывает на то что точка лежин внутри внешнего контура
                                                'блок для проверки условия
                                                Dim insBoolPoint As Boolean = False
                                                If IsNothing(PObl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolyline(PObl, Pt1.Convert2d(New Plane))
                                                ElseIf IsNothing(P2Obl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(P2Obl, Pt1)
                                                ElseIf IsNothing(CircObj) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(CircObj, Pt1)
                                                End If
                                                If insBoolPoint = True Then
                                                    Dim InsPoint As Point3d = New Point3d(XPt, YPt, 0)
                                                    If FormHatchCounter.CheckBox1.Checked = False Then
                                                        Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPoint, NameBlk, NameLayer, indColor, MScale, 20, Rot)
                                                        ids.Add(InsBlk.Id)
                                                    Else
                                                        Dim idmvbstyledict As ObjectId = mvbstyledict.GetAt(FormHatchCounter.ComboBox2.Text.Trim)
                                                        mvb = New Autodesk.Aec.DatabaseServices.MultiViewBlockReference()
                                                        mvb.BlockDefId = idmvbstyledict
                                                        If IsNothing(userSurface) = False Then
                                                            Try
                                                                Dim elev As Double = userSurface.FindElevationAtXY(XPt, YPt)
                                                                InsPoint = New Point3d(XPt, YPt, elev)
                                                                mvb.Location = InsPoint
                                                                mvb.Layer = obj.Layer
                                                                mvb.Scale = ScaleMVB
                                                                mvb.SetDatabaseDefaults()
                                                                acBlkTblRec.AppendEntity(mvb)
                                                                acTrans.AddNewlyCreatedDBObject(mvb, True)
                                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                            End Try
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Else
                                            If (i1 Mod 2) = 0 Then
                                                Dim XPt As Double = Math.Round(MinCoord.X + j1 * Shag, 3)
                                                Dim YPt As Double = Math.Round(MinCoord.Y + i1 * Shag, 3)
                                                Dim Pt1 As Point3d = New Point3d(XPt, YPt, 0)
                                                'флаг, который указывает на то что точка лежин внутри внешнего контура
                                                'блок для проверки условия
                                                Dim insBoolPoint As Boolean = False
                                                If IsNothing(PObl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolyline(PObl, Pt1.Convert2d(New Plane))
                                                ElseIf IsNothing(P2Obl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(P2Obl, Pt1)
                                                ElseIf IsNothing(CircObj) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(CircObj, Pt1)
                                                End If
                                                'проверяем блок для внутренних островков
                                                If insBoolPoint = True Then
                                                    Dim InsPoint As Point3d = New Point3d(XPt, YPt, 0)
                                                    If FormHatchCounter.CheckBox1.Checked = False Then
                                                        Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPoint, NameBlk, NameLayer, indColor, MScale, 20, Rot)
                                                        ids.Add(InsBlk.Id)
                                                    Else
                                                        Dim idmvbstyledict As ObjectId = mvbstyledict.GetAt(FormHatchCounter.ComboBox2.Text.Trim)
                                                        mvb = New Autodesk.Aec.DatabaseServices.MultiViewBlockReference()
                                                        mvb.BlockDefId = idmvbstyledict
                                                        If IsNothing(userSurface) = False Then
                                                            Try
                                                                Dim elev As Double = userSurface.FindElevationAtXY(XPt, YPt)
                                                                InsPoint = New Point3d(XPt, YPt, elev)
                                                                mvb.Location = InsPoint
                                                                mvb.Layer = obj.Layer
                                                                mvb.Scale = ScaleMVB
                                                                mvb.SetDatabaseDefaults()
                                                                acBlkTblRec.AppendEntity(mvb)
                                                                acTrans.AddNewlyCreatedDBObject(mvb, True)
                                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                            End Try
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i1
                                Next j1
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'штриховка в случайном порядке
                            Else
                                Dim randomObject As New Random() 'запускаем генератор случайных чисел
                                For j1 As Integer = 0 To i
                                    For i1 = 0 To j
                                        'генерируем случайные числа для вертикального и горизонтального смещения
                                        Dim randomNumberTop As Integer = randomObject.Next(0, 2)
                                        Dim randomNumberLeft As Integer = randomObject.Next(0, 2)
                                        Dim OffsetZn As Double = FormHatchCounter.NumericUpDown2.Value
                                        If (j1 Mod 2) = 0 Then
                                            If (i1 Mod 2) <> 0 Then
                                                Dim XPt As Double = Math.Round(MinCoord.X + j1 * Shag, 3)
                                                Dim YPt As Double = Math.Round(MinCoord.Y + i1 * Shag, 3)
                                                Dim Pt1 As Point3d = New Point3d(XPt, YPt, 0)
                                                'флаг, который указывает на то что точка лежин внутри внешнего контура
                                                'блок для проверки условия
                                                Dim insBoolPoint As Boolean = False
                                                If IsNothing(PObl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(PObl, Pt1)
                                                ElseIf IsNothing(P2Obl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(P2Obl, Pt1)
                                                ElseIf IsNothing(CircObj) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(CircObj, Pt1)
                                                End If
                                                'проверяем блок для внутренних островков
                                                If insBoolPoint = True Then 'если точка расположена внутри внешнего контура тогда
                                                    'определяем величину смещения
                                                    If insBoolPoint = True Then
                                                        If randomNumberTop = 1 Then
                                                            YPt = YPt - OffsetZn
                                                        ElseIf randomNumberTop = 2 Then
                                                            YPt = YPt + OffsetZn
                                                        Else
                                                            YPt = YPt
                                                        End If
                                                        If randomNumberLeft = 1 Then
                                                            XPt = XPt - OffsetZn
                                                        ElseIf randomNumberLeft = 2 Then
                                                            XPt = XPt + OffsetZn
                                                        Else
                                                            XPt = XPt
                                                        End If
                                                    End If
                                                    Dim InsPoint As Point3d = New Point3d(XPt, YPt, 0)
                                                    If FormHatchCounter.CheckBox1.Checked = False Then
                                                        Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPoint, NameBlk, NameLayer, indColor, MScale, 20, Rot)
                                                        ids.Add(InsBlk.Id)
                                                    Else
                                                        Dim idmvbstyledict As ObjectId = mvbstyledict.GetAt(FormHatchCounter.ComboBox2.Text.Trim)
                                                        mvb = New Autodesk.Aec.DatabaseServices.MultiViewBlockReference()
                                                        mvb.BlockDefId = idmvbstyledict
                                                        If IsNothing(userSurface) = False Then
                                                            Try
                                                                Dim elev As Double = userSurface.FindElevationAtXY(XPt, YPt)
                                                                InsPoint = New Point3d(XPt, YPt, elev)
                                                                mvb.Location = InsPoint
                                                                mvb.Layer = obj.Layer
                                                                mvb.Scale = ScaleMVB
                                                                mvb.SetDatabaseDefaults()
                                                                acBlkTblRec.AppendEntity(mvb)
                                                                acTrans.AddNewlyCreatedDBObject(mvb, True)
                                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                            End Try
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Else
                                            If (i1 Mod 2) = 0 Then
                                                Dim XPt As Double = Math.Round(MinCoord.X + j1 * Shag, 3)
                                                Dim YPt As Double = Math.Round(MinCoord.Y + i1 * Shag, 3)
                                                Dim Pt1 As Point3d = New Point3d(XPt, YPt, 0)
                                                'флаг, который указывает на то что точка лежин внутри внешнего контура
                                                'блок для проверки условия
                                                Dim insBoolPoint As Boolean = False
                                                If IsNothing(PObl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(PObl, Pt1)
                                                ElseIf IsNothing(P2Obl) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(P2Obl, Pt1)
                                                ElseIf IsNothing(CircObj) = False Then
                                                    insBoolPoint = MathFunction.FuncCrossPointPolygon(CircObj, Pt1)
                                                End If
                                                If insBoolPoint = True Then 'если точка расположена внутри внешнего контура тогда
                                                    'проверяем на наличие точки во внутреннем островке
                                                    'определяем величину смещения
                                                    If insBoolPoint = True Then
                                                        If randomNumberTop = 1 Then
                                                            YPt = YPt - OffsetZn
                                                        ElseIf randomNumberTop = 2 Then
                                                            YPt = YPt + OffsetZn
                                                        Else
                                                            YPt = YPt
                                                        End If
                                                        If randomNumberLeft = 1 Then
                                                            XPt = XPt - OffsetZn
                                                        ElseIf randomNumberLeft = 2 Then
                                                            XPt = XPt + OffsetZn
                                                        Else
                                                            XPt = XPt
                                                        End If
                                                    End If
                                                    Dim InsPoint As Point3d = New Point3d(XPt, YPt, 0)
                                                    If FormHatchCounter.CheckBox1.Checked = False Then
                                                        Dim InsBlk As BlockReference = AcadFunction.FuncInsertDrawingBlk(InsPoint, NameBlk, NameLayer, indColor, MScale, 20, Rot)
                                                        ids.Add(InsBlk.Id)
                                                    Else
                                                        Dim idmvbstyledict As ObjectId = mvbstyledict.GetAt(FormHatchCounter.ComboBox2.Text.Trim)
                                                        mvb = New Autodesk.Aec.DatabaseServices.MultiViewBlockReference()
                                                        mvb.BlockDefId = idmvbstyledict
                                                        If IsNothing(userSurface) = False Then
                                                            Try
                                                                Dim elev As Double = userSurface.FindElevationAtXY(XPt, YPt)
                                                                InsPoint = New Point3d(XPt, YPt, elev)
                                                                mvb.Location = InsPoint
                                                                mvb.Layer = obj.Layer
                                                                mvb.Scale = ScaleMVB
                                                                mvb.SetDatabaseDefaults()
                                                                acBlkTblRec.AppendEntity(mvb)
                                                                acTrans.AddNewlyCreatedDBObject(mvb, True)
                                                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                                            End Try
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i1
                                Next j1
                            End If
                            If ids.Count > 0 Then
                                Dim RNG As Random = New Random
                                Dim Numb As Integer = RNG.Next
                                Dim grp As Group = New Group("TOPO" & Numb, True)
                                gd.UpgradeOpen()
                                SymbolUtilityServices.ValidateSymbolName("q", False)
                                Dim grpId As ObjectId = gd.SetAt("q", grp)
                                acTrans.AddNewlyCreatedDBObject(grp, True)
                                grp.InsertAt(0, ids)
                            End If
                        End If
                    End If
                Next
            End If
            acTrans.Commit()
        End Using
    End Sub
End Class
