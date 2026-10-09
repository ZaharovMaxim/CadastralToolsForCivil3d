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

Public Class WireSagAirLine
    'рисование проводов
    <CommandMethod("WireSagAirLine")>
    Public Sub WireSagAirLine()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acDocEd As Editor = acDoc.Editor
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
        '=========================================================================================================
        'запускаем форму
        Dim FormWire As FormWireSagAirLine = New FormWireSagAirLine
        FormWire.ShowDialog()
        If FormWire.boolFormShow = False Then
            Exit Sub
        End If

        Dim shInterpol As Integer = FormWire.NumericUpDown2.Value
        'Массив для вершин линий
        Dim ArrayPointVert As Double(,) = Nothing
        Dim CountArrayPointVert As Integer = 0
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'рабочая поверхность
            Dim TinSurf As TinSurface = Nothing
            Dim boolSurf As Boolean = False
            Dim NameSurf As String = FormWire.ComboBox1.Text.Trim
            If NameSurf.Length > 0 Then
                For Each idSurf As ObjectId In CivilDoc.GetSurfaceIds
                    TinSurf = tr.GetObject(idSurf, OpenMode.ForRead)
                    If TinSurf.Name.Trim Like NameSurf Then
                        boolSurf = True
                        Exit For
                    End If
                Next
            End If
            'величина провиса
            '================================================================================================
            Dim h As Double = FormWire.NumericUpDown1.Value
            Dim hPr As Double = h 'величина провиса
            If h = 0 Then
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите величину провиса: ")
                pStrOpts.DefaultValue = 8
                Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
                h = Val(pStrRes.Value)
                hPr = h
            End If
            'выбираем 3d полилинию или Характерную линию
            '================================================================================================
            Dim resBlk1 As PromptEntityResult = Nothing
            Dim opt1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите Характерную линию: ")
            resBlk1 = acDocEd.GetEntity(opt1)
            If Not (resBlk1.Status = PromptStatus.OK) Then
                Exit Sub
            End If
            Dim pvId1 As ObjectId = resBlk1.ObjectId
            Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pvId1, OpenMode.ForWrite)
            Dim PtSelect As Point3d = resBlk1.PickedPoint
            Dim FeatLine As FeatureLine = Nothing
            Dim acPolyLine3d As Polyline3d = Nothing
            Dim boolFeatLine As Boolean = False
            Dim boolPolyline3d As Boolean = False
            If TypeOf obj Is FeatureLine Then
                FeatLine = obj
                Dim PointColl As Point3dCollection = New Point3dCollection
                Dim ftype As Autodesk.Civil.FeatureLinePointType = Autodesk.Civil.FeatureLinePointType.PIPoint
                PointColl = FeatLine.GetPoints(ftype)
                If PointColl.Count > 0 Then
                    For Each pt3d As Point3d In PointColl
                        ReDim Preserve ArrayPointVert(2, CountArrayPointVert)
                        ArrayPointVert(0, CountArrayPointVert) = pt3d.X
                        ArrayPointVert(1, CountArrayPointVert) = pt3d.Y
                        ArrayPointVert(2, CountArrayPointVert) = pt3d.Z
                        CountArrayPointVert += 1
                    Next
                End If
                boolFeatLine = True
            ElseIf TypeOf obj Is Polyline3d Then
                acPolyLine3d = obj
                For Each idVert As ObjectId In acPolyLine3d
                    Dim vert As PolylineVertex3d = tr.GetObject(idVert, OpenMode.ForRead)
                    If vert.IsErased = False Then
                        ReDim Preserve ArrayPointVert(2, CountArrayPointVert)
                        ArrayPointVert(0, CountArrayPointVert) = vert.Position.X
                        ArrayPointVert(1, CountArrayPointVert) = vert.Position.Y
                        ArrayPointVert(2, CountArrayPointVert) = vert.Position.Z
                        CountArrayPointVert += 1
                    End If
                Next
                boolPolyline3d = True
            End If
            '===========================================================================================
            If IsArray(ArrayPointVert) = False Then
                Exit Sub
            End If
            If ArrayPointVert.GetUpperBound(1) < 1 Then
                Exit Sub
            End If
            '===============================================================================================
            Dim indStartPoint As Integer = 0
            Dim indEndPoint As Integer = 1
            'смотрим количество вершин
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Dim PicPoint As Point3d = Nothing
            Dim hground As Double = -99999999
            'запрос на обработку одного сегмента
            If FormWire.RadioButton5.Checked = True Then
                Dim PointWireSelect As Point2d = FormWire.PointWire
                If PointWireSelect.X = 0 And PointWireSelect.Y = 0 Then
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.Message = vbLf & "Укажите точку замера провиса: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.OK Then
                        PointWireSelect = pPtRes.Value.Convert2d(New Plane)
                    Else
                        Exit Sub
                    End If
                End If
                If boolSurf = True Then
                    If FormWire.RadioButton1.Checked = True Then
                        Try
                            hground = TinSurf.FindElevationAtXY(PointWireSelect.X, PointWireSelect.Y)
                            h = h + hground
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            MsgBox("Выбранная точка находится за пределами рабочей поверхности. Макро прервано!")
                            Exit Sub
                        End Try
                    End If
                End If
                '===================================================================
                'точка провиса
                PicPoint = New Point3d(PointWireSelect.X, PointWireSelect.Y, h)
                '=====================================================================
                'вспомогательная линия
                Dim acLine As Line = Nothing
                Dim boolFindPicPoint As Boolean = False
                Dim MinLine As Double = 10000000
                For i As Integer = 0 To ArrayPointVert.GetUpperBound(1) - 1
                    Dim pt1 As Point3d = New Point3d(ArrayPointVert(0, i), ArrayPointVert(1, i), ArrayPointVert(2, i))
                    Dim pt2 As Point3d = New Point3d(ArrayPointVert(0, i + 1), ArrayPointVert(1, i + 1), ArrayPointVert(2, i + 1))
                    Dim acLineTemp As Line = New Line(pt1, pt2)
                    Try
                        Dim Temppoint As Point3d = acLineTemp.GetClosestPointTo(PicPoint, False)
                        If Temppoint.Convert2d(New Plane).GetVectorTo(PicPoint.Convert2d(New Plane)).Length < MinLine Then
                            MinLine = Temppoint.Convert2d(New Plane).GetVectorTo(PicPoint.Convert2d(New Plane)).Length
                            boolFindPicPoint = True
                            acLine = acLineTemp
                        End If
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                    End Try
                Next i
                If boolFindPicPoint = False Then
                    MsgBox("Не удалось определить сегмент Линии на которой был определен провил. Укажите точку на расстоянии не менее 0.5м от нужного сегмента!!!")
                    Exit Sub
                End If
                If IsNothing(acLine) = False Then
                    If acLine.Length > 0 Then
                        '=============================================================================
                        'ищем расстояние от начала выбранного сегмента до точки провиса
                        Dim TempSenterPoint As Point3d = New Point3d((acLine.StartPoint.X + acLine.EndPoint.X) / 2, (acLine.StartPoint.Y + acLine.EndPoint.Y) / 2, (acLine.StartPoint.Z + acLine.EndPoint.Z) / 2)
                        Dim DistTempPicPoint As Double = acLine.Length / 2
                        Try
                            TempSenterPoint = acLine.GetClosestPointTo(PicPoint, False)
                            DistTempPicPoint = acLine.GetDistAtPoint(TempSenterPoint)
                        Catch ex As Autodesk.AutoCAD.Runtime.Exception
                            MsgBox("Не удалось найти точку замера провиса нижнего провода.")
                        End Try

                        If TempSenterPoint.Z - h < 0 Then
                            MsgBox("Большая погрешность определения высот провиса провода. Точка замера провиса выше точек замера подвеса проводов на величину: " & TempSenterPoint.Z - h & "Проверьте исходные данные (отметки подвесов проводов и величину провиса провода!!!")
                            Exit Sub
                        End If

                        Dim boolInsWire As Boolean = False
                        Dim HPic As Double = 0
                        'ищем центр дуги
                        Dim TempPt1 As Point2d = New Point2d(0, acLine.StartPoint.Z)
                        Dim v1 As Double = acLine.StartPoint.Convert2d(New Plane).GetVectorTo(acLine.EndPoint.Convert2d(New Plane)).Length
                        Dim TempPt2 As Point2d = New Point2d(v1, acLine.EndPoint.Z)
                        Dim v2 As Double = acLine.StartPoint.Convert2d(New Plane).GetVectorTo(PicPoint.Convert2d(New Plane)).Length
                        Dim TempPt3 As Point2d = New Point2d(v2, PicPoint.Z)
                        '===========================================================================================
                        'ищем координаты центра дуги
                        Dim CenterArc As Point2d = MathFunction.FuncFindMiddleArc(TempPt1, TempPt2, TempPt3)
                        Dim CenterPoint As Point2d = MathFunction.FuncFindCenterCircleByThreePoint(New Point3d(TempPt1.X, TempPt1.Y, 0), New Point3d(TempPt3.X, TempPt3.Y, 0), New Point3d(TempPt2.X, TempPt2.Y, 0))
                        Dim CircArc As CircularArc3d = New CircularArc3d(New Point3d(TempPt1.X, TempPt1.Y, 0), New Point3d(TempPt3.X, TempPt3.Y, 0), New Point3d(TempPt2.X, TempPt2.Y, 0))
                        Dim StAngle As Double = CenterPoint.GetVectorTo(TempPt1).Angle
                        Dim EnAngle As Double = CenterPoint.GetVectorTo(TempPt2).Angle
                        Dim r As Double = CenterPoint.GetVectorTo(TempPt1).Length
                        Dim ArcTemp As Arc = New Arc(New Point3d(CenterPoint.X, CenterPoint.Y, 0), r, StAngle, EnAngle)
                        If ArcTemp.Length > 1 Then
                            For i2 As Integer = 1 To ArcTemp.Length Step shInterpol
                                If i2 >= ArcTemp.Length Then
                                    Exit For
                                End If
                                If boolInsWire = False Then
                                    If i2 = DistTempPicPoint Then
                                        boolInsWire = True
                                    ElseIf i2 > DistTempPicPoint Then
                                        If boolFeatLine = True Then
                                            Try
                                                Dim TempPtPline As Point3d = ArcTemp.GetPointAtDist(DistTempPicPoint)
                                                Dim PMTP As Point2d = New Point2d(TempPtPline.X, TempPtPline.Z)
                                                Dim PtTempXYZ As Point3d = acLine.GetPointAtDist(DistTempPicPoint)
                                                FeatLine.InsertElevationPoint(New Point3d(PtTempXYZ.X, PtTempXYZ.Y, TempPtPline.Y))
                                                boolInsWire = True
                                                If FormWire.RadioButton1.Checked = True Then
                                                    HPic = TempPtPline.Y - hground - hPr
                                                Else
                                                    HPic = TempPtPline.Y - h
                                                End If
                                            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                            End Try
                                        End If
                                    End If
                                End If
                                Try
                                    Dim TempPtPline As Point3d = ArcTemp.GetPointAtDist(i2)
                                    Dim PMTP As Point2d = New Point2d(TempPtPline.X, TempPtPline.Z)
                                    Dim PtTempXYZ As Point3d = acLine.GetPointAtDist(i2)
                                    If boolFeatLine = True Then
                                        FeatLine.InsertElevationPoint(New Point3d(PtTempXYZ.X, PtTempXYZ.Y, TempPtPline.Y))
                                    End If
                                Catch ex As System.ArgumentException
                                Catch ex1 As Autodesk.AutoCAD.Runtime.Exception
                                Catch ex1 As System.Exception
                                End Try
                            Next i2
                        End If
                        MsgBox("Погрешность интерполяции дуги составила: " & Math.Round(HPic, 3) & "м.")
                    End If
                End If
            ElseIf FormWire.RadioButton3.Checked = True Then
                For i As Integer = 0 To ArrayPointVert.GetUpperBound(1) - 1
                    Dim Pt1 As Point3d = New Point3d(ArrayPointVert(0, i), ArrayPointVert(1, i), ArrayPointVert(2, i))
                    Dim Pt2 As Point3d = New Point3d(ArrayPointVert(0, i + 1), ArrayPointVert(1, i + 1), ArrayPointVert(2, i + 1))
                    Dim acLine As Line = New Line(Pt1, Pt2)
                    If boolSurf = True Then
                        If FormWire.RadioButton1.Checked = True Then
                            Try
                                hground = TinSurf.FindElevationAtXY((Pt1.X + Pt2.X) / 2, (Pt1.Y + Pt2.Y) / 2)
                                h = h + hground
                                PicPoint = New Point3d((Pt1.X + Pt2.X) / 2, (Pt1.Y + Pt2.Y) / 2, h)
                            Catch ex As Autodesk.Civil.PointNotOnEntityException
                                Continue For
                            End Try
                        End If
                    End If
                    If acLine.Length > 0 Then
                        'ищем центр дуги
                        Dim TempPt1 As Point2d = New Point2d(0, acLine.StartPoint.Z)
                        Dim v1 As Double = acLine.StartPoint.Convert2d(New Plane).GetVectorTo(acLine.EndPoint.Convert2d(New Plane)).Length
                        Dim TempPt2 As Point2d = New Point2d(v1, acLine.EndPoint.Z)
                        Dim v2 As Double = acLine.StartPoint.Convert2d(New Plane).GetVectorTo(PicPoint.Convert2d(New Plane)).Length
                        Dim TempPt3 As Point2d = New Point2d(v2, PicPoint.Z)
                        '===========================================================================================
                        'ищем координаты центра дуги
                        Dim CenterArc As Point2d = MathFunction.FuncFindMiddleArc(TempPt1, TempPt2, TempPt3)
                        Dim CenterPoint As Point2d = MathFunction.FuncFindCenterCircleByThreePoint(New Point3d(TempPt1.X, TempPt1.Y, 0), New Point3d(TempPt3.X, TempPt3.Y, 0), New Point3d(TempPt2.X, TempPt2.Y, 0))
                        'Dim CircArc As CircularArc3d = New CircularArc3d(acLine.StartPoint, PicPoint, acLine.EndPoint)
                        Dim CircArc As CircularArc3d = New CircularArc3d(New Point3d(TempPt1.X, TempPt1.Y, 0), New Point3d(TempPt3.X, TempPt3.Y, 0), New Point3d(TempPt2.X, TempPt2.Y, 0))
                        Dim StAngle As Double = CenterPoint.GetVectorTo(TempPt1).Angle
                        Dim EnAngle As Double = CenterPoint.GetVectorTo(TempPt2).Angle
                        Dim r As Double = CenterPoint.GetVectorTo(TempPt1).Length
                        Dim ArcTemp As Arc = New Arc(New Point3d(CenterPoint.X, CenterPoint.Y, 0), r, StAngle, EnAngle)
                        If ArcTemp.Length > 1 Then
                            For i2 As Integer = 1 To ArcTemp.Length Step shInterpol
                                If i2 >= ArcTemp.Length Then
                                    Exit For
                                End If
                                Try
                                    Dim TempPtPline As Point3d = ArcTemp.GetPointAtDist(i2)
                                    Dim PMTP As Point2d = New Point2d(TempPtPline.X, TempPtPline.Z)
                                    Dim PtTempXYZ As Point3d = acLine.GetPointAtDist(i2)
                                    If boolFeatLine = True Then
                                        FeatLine.InsertElevationPoint(New Point3d(PtTempXYZ.X, PtTempXYZ.Y, TempPtPline.Y))
                                    End If
                                Catch ex As System.ArgumentException
                                Catch ex1 As Autodesk.AutoCAD.Runtime.Exception
                                Catch ex1 As System.Exception
                                End Try
                            Next i2
                        End If
                    End If
                Next i
            End If
            tr.Commit()
        End Using
    End Sub
End Class
