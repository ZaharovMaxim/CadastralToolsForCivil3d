Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class DrawSlopes
    '==========================================================================
    'функция рисует неукрепленный откос (новый вариант)
    Public Shared Function FuncOtkosDraw(ByVal ArrayPoint As Double(,), ByVal brv As Curve, ByVal pod As Curve, ByVal TipOtkos As String, ByVal LenSh As Double, Optional ByVal nameLayer As String = "0", Optional ByVal nameColor As String = "7", Optional ByVal guidOtkos As String = "0000", Optional ByVal HandleBrv As String = "", Optional ByVal m As Integer = 500, Optional ByVal ShBool As Boolean = False, Optional ByVal WidthSh As Double = 20) As Boolean
        FuncOtkosDraw = False
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'определяем тип откоса
            If TipOtkos Like "Откос неукрепленный" Then
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                    Dim PointBrv1 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i)) 'начало длинного штриха (бровка)
                    Dim PointPod1 As Point2d = New Point2d(ArrayPoint(2, i), ArrayPoint(3, i)) 'конец длинного штриха (подошва)
                    If PointPod1.X = 0 And PointPod1.Y = 0 Then
                        GoTo line1
                    End If
                    'строим длинный штрих
                    Dim acPLine As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                    acPLine.SetDatabaseDefaults()
                    acPLine.AddVertexAt(0, PointBrv1, 0, 0, 0)
                    acPLine.AddVertexAt(1, PointPod1, 0, 0, 0)
                    acPLine.Layer = nameLayer
                    acPLine.ColorIndex = nameColor
                    acPLine.LineWeight = WidthSh
                    acBlkTblRec.AppendEntity(acPLine)
                    acTrans.AddNewlyCreatedDBObject(acPLine, True)
                    'добавляем словарь в примитив
                    Dim extId As ObjectId = acPLine.ExtensionDictionary
                    If extId.IsNull = True Then
                        acPLine.UpgradeOpen()
                        acPLine.CreateExtensionDictionary()
                        extId = acPLine.ExtensionDictionary
                    End If
                    Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                    dbExt.UpgradeOpen()
                    Dim xRec As Xrecord = New Xrecord()
                    Dim rb As ResultBuffer = New ResultBuffer()
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Длинный штрих")) 'тип штриха
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                    rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                    xRec.Data = rb
                    dbExt.SetAt("SLOPEUSER", xRec)
                    acTrans.AddNewlyCreatedDBObject(xRec, True)
                    If i < ArrayPoint.GetUpperBound(1) Then
                        'второй длинный штрих (определяем координаты)
                        Dim PointBrv2 As Point3d = New Point3d(ArrayPoint(0, i + 1), ArrayPoint(1, i + 1), 0) 'начало длинного штриха (бровка)
                        Dim PointPod2 As Point3d = New Point3d(ArrayPoint(2, i + 1), ArrayPoint(3, i + 1), 0) 'конец длинного штриха (подошва)
                        'определяем координаты короткого штриха
                        '1. определям дистанцию до штрихов на бровке
                        Dim PointShBrv As Point3d = Nothing
                        Try
                            Dim PointBrv13d As Point3d = New Point3d(ArrayPoint(0, i), ArrayPoint(1, i), 0)
                            Dim DistBrv1 As Double = brv.GetDistAtPoint(PointBrv13d)
                            Dim DistBrv2 As Double = brv.GetDistAtPoint(PointBrv2)
                            Dim dLenBrv As Double = (DistBrv1 + DistBrv2) / 2
                            PointShBrv = brv.GetPointAtDist(dLenBrv)
                        Catch ex As Exception
                            PointShBrv = New Point3d((PointBrv1.X + PointBrv2.X) / 2, (PointBrv1.Y + PointBrv2.Y) / 2, 0)
                        End Try
                        '2. определяем дистанцию по подошве
                        Dim PointShPod As Point3d = Nothing
                        Try
                            Dim TempPointPod13d As Point3d = New Point3d(ArrayPoint(2, i), ArrayPoint(3, i), 0)
                            Dim PointPod13d As Point3d = pod.GetClosestPointTo(TempPointPod13d, True)
                            Dim DistPod1 As Double = pod.GetDistAtPoint(PointPod13d)
                            If PointPod2.X = 0 And PointPod2.Y = 0 Then
                                GoTo line1
                            End If
                            Dim DistPod2 As Double = pod.GetDistAtPoint(PointPod2)
                            Dim dLenPod As Double = (DistPod1 + DistPod2) / 2
                            PointShPod = pod.GetPointAtDist(dLenPod)
                        Catch ex As Exception
                            PointShPod = New Point3d((PointPod1.X + PointPod2.X) / 2, (PointPod1.Y + PointPod2.Y) / 2, 0)
                        End Try
                        '2. строим короткий штрих
                        Dim PtShBrv As Point2d = New Point2d(PointShBrv.X, PointShBrv.Y)
                        Dim PtShPod As Point2d = New Point2d(PointShPod.X, PointShPod.Y)
                        Dim Vect As Vector2d = PtShBrv.GetVectorTo(PtShPod)
                        Dim LenghtVect As Double = Vect.Length
                        'смотрим длину короткого штриха
                        If ShBool = True Then
                            LenghtVect = LenghtVect / 2
                        Else
                            Dim Lsr As Double = LenSh + LenSh * 0.25
                            If LenghtVect < Lsr Then
                                LenghtVect = LenghtVect * 0.5
                            Else
                                LenghtVect = LenSh
                            End If
                        End If
                        Dim xSh As Double = PtShBrv.X + LenghtVect * Math.Cos(Vect.Angle)
                        Dim YSh As Double = PtShBrv.Y + LenghtVect * Math.Sin(Vect.Angle)
                        Dim PtSh As Point2d = New Point2d(xSh, YSh)
                        Dim acPLineSh As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                        acPLineSh.SetDatabaseDefaults()
                        acPLineSh.AddVertexAt(0, PtShBrv, 0, 0, 0)
                        acPLineSh.AddVertexAt(1, PtSh, 0, 0, 0)
                        acPLineSh.Layer = nameLayer
                        acPLineSh.ColorIndex = nameColor
                        acPLineSh.LineWeight = WidthSh
                        acBlkTblRec.AppendEntity(acPLineSh)
                        acTrans.AddNewlyCreatedDBObject(acPLineSh, True)
                        'добавляем словарь в примитив
                        extId = acPLineSh.ExtensionDictionary
                        If extId.IsNull = True Then
                            acPLineSh.UpgradeOpen()
                            acPLineSh.CreateExtensionDictionary()
                            extId = acPLineSh.ExtensionDictionary
                        End If
                        dbExt = acTrans.GetObject(extId, OpenMode.ForRead)
                        dbExt.UpgradeOpen()
                        xRec = New Xrecord()
                        rb = New ResultBuffer()
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Короткий штрих")) 'тип штриха
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                        rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                        xRec.Data = rb
                        dbExt.SetAt("SLOPEUSER", xRec)
                        acTrans.AddNewlyCreatedDBObject(xRec, True)
                    End If
line1:
                Next
            ElseIf TipOtkos Like "Откос укрепленный" Then
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1) - 1
                    Dim PointBrv1 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i)) 'начало длинного штриха (бровка)
                    Dim PointPod1 As Point2d = New Point2d(ArrayPoint(2, i), ArrayPoint(3, i)) 'конец длинного штриха (подошва)
                    If PointPod1.X = 0 And PointPod1.Y = 0 Then
                        GoTo line2
                    End If
                    'строим длинный штрих
                    Dim acPLine As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                    acPLine.SetDatabaseDefaults()
                    acPLine.AddVertexAt(0, PointBrv1, 0, 0, 0)
                    acPLine.AddVertexAt(1, PointPod1, 0, 0, 0)
                    acPLine.Layer = nameLayer
                    acPLine.ColorIndex = nameColor
                    acPLine.LineWeight = WidthSh
                    acBlkTblRec.AppendEntity(acPLine)
                    acTrans.AddNewlyCreatedDBObject(acPLine, True)
                    If acPLine.Id.IsNull = False Then
                        'добавляем словарь в примитив
                        Dim extId As ObjectId = acPLine.ExtensionDictionary
                        If extId.IsNull = True Then
                            acPLine.UpgradeOpen()
                            acPLine.CreateExtensionDictionary()
                            extId = acPLine.ExtensionDictionary
                        End If
                        Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                        dbExt.UpgradeOpen()
                        Dim xRec As Xrecord = New Xrecord()
                        Dim rb As ResultBuffer = New ResultBuffer()
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Длинный штрих")) 'тип штриха
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                        rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                        xRec.Data = rb
                        dbExt.SetAt("SLOPEUSER", xRec)
                        acTrans.AddNewlyCreatedDBObject(xRec, True)
                    End If
                    '============================================================================================================
                    'рисуем короткий штрих
                    'смотрим второй длинный штрих
                    If i < ArrayPoint.GetUpperBound(1) Then
                        Dim PointBrv2 As Point2d = New Point2d(ArrayPoint(0, i + 1), ArrayPoint(1, i + 1)) 'начало длинного штриха (бровка)
                        Dim PointPod2 As Point2d = New Point2d(ArrayPoint(2, i + 1), ArrayPoint(3, i + 1)) 'конец длинного штриха (подошва)
                        Dim PointBrv13d As Point3d = New Point3d(ArrayPoint(0, i), ArrayPoint(1, i), 0)
                        Dim PointPod13d As Point3d = New Point3d(ArrayPoint(2, i), ArrayPoint(3, i), 0)
                        Dim PointBrv23d As Point3d = New Point3d(ArrayPoint(0, i + 1), ArrayPoint(1, i + 1), 0)
                        Dim PointPod23d As Point3d = New Point3d(ArrayPoint(2, i + 1), ArrayPoint(3, i + 1), 0)
                        'определяем координаты короткого штриха
                        '1. определям дистанцию до штрихов на бровке
                        Dim PointShBrv As Point3d = Nothing
                        Try
                            Dim DistBrv1 As Double = brv.GetDistAtPoint(PointBrv13d)
                            Dim DistBrv2 As Double = brv.GetDistAtPoint(PointBrv23d)

                            Dim dLenBrv As Double = (DistBrv1 + DistBrv2) / 2
                            PointShBrv = brv.GetPointAtDist(dLenBrv)
                        Catch ex As Exception
                            PointShBrv = New Point3d((PointBrv1.X + PointBrv2.X) / 2, (PointBrv1.Y + PointBrv2.Y) / 2, 0)
                        End Try
                        '2. определяем дистанцию по подошве
                        Dim PointShPod As Point3d = Nothing
                        Try
                            Dim DistPod1 As Double = pod.GetDistAtPoint(PointPod13d)
                            If PointPod2.X = 0 And PointPod2.Y = 0 Then
                                GoTo line2
                            End If
                            Dim DistPod2 As Double = pod.GetDistAtPoint(PointPod23d)
                            Dim dLenPod As Double = (DistPod1 + DistPod2) / 2
                            PointShPod = pod.GetPointAtDist(dLenPod)
                        Catch ex As Exception
                            PointShPod = New Point3d((PointPod1.X + PointPod2.X) / 2, (PointPod1.Y + PointPod2.Y) / 2, 0)
                        End Try
                        '2. строим короткий штрих
                        Dim PtShBrv As Point2d = New Point2d(PointShBrv.X, PointShBrv.Y)
                        Dim PtShPod As Point2d = New Point2d(PointShPod.X, PointShPod.Y)
                        Dim Vect As Vector2d = PtShBrv.GetVectorTo(New Point2d(PointShPod.X, PointShPod.Y))
                        Dim LenghtVect As Double = Vect.Length
                        'смотрим длину короткого штриха
                        If ShBool = True Then
                            LenghtVect = LenghtVect / 2
                        Else
                            Dim Lsr As Double = LenSh + LenSh * 0.25
                            If LenghtVect < Lsr Then
                                LenghtVect = LenghtVect * 0.5
                            Else
                                LenghtVect = LenSh
                            End If
                        End If
                        'проверяем место для точки
                        Dim xSh As Double = PtShBrv.X + LenghtVect * Math.Cos(Vect.Angle)
                        Dim YSh As Double = PtShBrv.Y + LenghtVect * Math.Sin(Vect.Angle)
                        Dim PtSh As Point2d = New Point2d(xSh, YSh)
                        Dim acPLineSh As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                        acPLineSh.SetDatabaseDefaults()
                        acPLineSh.AddVertexAt(0, PtShBrv, 0, 0, 0)
                        acPLineSh.AddVertexAt(1, PtSh, 0, 0, 0)
                        acPLineSh.Layer = nameLayer
                        acPLineSh.ColorIndex = nameColor
                        acPLineSh.LineWeight = WidthSh
                        acBlkTblRec.AppendEntity(acPLineSh)
                        acTrans.AddNewlyCreatedDBObject(acPLineSh, True)
                        If acPLine.ObjectId.IsNull = False Then
                            Dim extId As ObjectId = acPLineSh.ExtensionDictionary
                            If extId.IsNull = True Then
                                acPLineSh.UpgradeOpen()
                                acPLineSh.CreateExtensionDictionary()
                                extId = acPLineSh.ExtensionDictionary
                            End If
                            Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                            dbExt.UpgradeOpen()
                            Dim xRec As Xrecord = New Xrecord()
                            Dim rb As ResultBuffer = New ResultBuffer()
                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Короткий штрих")) 'тип штриха
                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                            rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                            xRec.Data = rb
                            dbExt.SetAt("SLOPEUSER", xRec)
                            acTrans.AddNewlyCreatedDBObject(xRec, True)
                        End If
                        'строим точку
                        If LenghtVect - 0.35 > 0 Then
                            Dim l1 As Double = LenghtVect + 0.25
                            Dim dx As Double = l1 * Math.Cos(Vect.Angle)
                            Dim dy As Double = l1 * Math.Sin(Vect.Angle)
                            Dim PointStart As Point2d = New Point2d(PtShBrv.X + dx, PtShBrv.Y + dy) 'расчет начальной точки
                            Dim l2 As Double = LenghtVect + 0.35
                            dx = l2 * Math.Cos(Vect.Angle)
                            dy = l2 * Math.Sin(Vect.Angle)
                            Dim PointEnd As Point2d = New Point2d(PtShBrv.X + dx, PtShBrv.Y + dy) 'расчет конечной точки
                            Dim acPLinePoint As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                            acPLinePoint.SetDatabaseDefaults()
                            acPLinePoint.AddVertexAt(0, PointStart, 0, 0, 0)
                            acPLinePoint.AddVertexAt(1, PointEnd, 0, 0, 0)
                            acPLinePoint.Layer = nameLayer
                            acPLinePoint.ColorIndex = nameColor
                            acPLinePoint.LineWeight = WidthSh
                            acBlkTblRec.AppendEntity(acPLinePoint)
                            acTrans.AddNewlyCreatedDBObject(acPLinePoint, True)
                            'добавляем словарь в примитив
                            If acPLinePoint.Id.IsNull = False Then
                                Dim extId As ObjectId = acPLinePoint.ExtensionDictionary
                                If extId.IsNull = True Then
                                    acPLinePoint.UpgradeOpen()
                                    acPLinePoint.CreateExtensionDictionary()
                                    extId = acPLinePoint.ExtensionDictionary
                                End If
                                Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                                dbExt.UpgradeOpen()
                                Dim xRec As Xrecord = New Xrecord()
                                Dim rb As ResultBuffer = New ResultBuffer()
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Точка")) 'тип штриха
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                                rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                                rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                                xRec.Data = rb
                                dbExt.SetAt("SLOPEUSER", xRec)
                                acTrans.AddNewlyCreatedDBObject(xRec, True)
                            End If
                        End If
                    End If
line2:
                Next
            ElseIf TipOtkos Like "Обрыв земляной" Then
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                    Dim PointBrv1 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i)) 'начало длинного штриха (бровка)
                    Dim PointPod1 As Point2d = New Point2d(ArrayPoint(2, i), ArrayPoint(3, i)) 'конец длинного штриха (подошва)
                    If PointPod1.X = 0 And PointPod1.Y = 0 Then
                        GoTo line3
                    End If
                    'строим длинный штрих
                    Dim acPLine As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                    acPLine.SetDatabaseDefaults()
                    acPLine.AddVertexAt(0, PointBrv1, 0, 0, 0)
                    acPLine.AddVertexAt(1, PointPod1, 0, 0, 0)
                    acPLine.Layer = nameLayer
                    acPLine.ColorIndex = nameColor
                    acPLine.LineWeight = WidthSh
                    acBlkTblRec.AppendEntity(acPLine)
                    acTrans.AddNewlyCreatedDBObject(acPLine, True)
                    If acPLine.Id.IsNull = False Then
                        'добавляем словарь в примитив
                        Dim extId As ObjectId = acPLine.ExtensionDictionary
                        If extId.IsNull = True Then
                            acPLine.UpgradeOpen()
                            acPLine.CreateExtensionDictionary()
                            extId = acPLine.ExtensionDictionary
                        End If
                        Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                        dbExt.UpgradeOpen()
                        Dim xRec As Xrecord = New Xrecord()
                        Dim rb As ResultBuffer = New ResultBuffer()
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Длинный штрих")) 'тип штриха
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                        rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                        rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                        xRec.Data = rb
                        dbExt.SetAt("SLOPEUSER", xRec)
                        acTrans.AddNewlyCreatedDBObject(xRec, True)
                    End If
line3:
                Next
            ElseIf TipOtkos Like "Стенка подпорная каменная" Then
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                    Dim PointBrv1 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i)) 'начало длинного штриха (бровка)
                    Dim PointPod1 As Point2d = New Point2d(ArrayPoint(2, i), ArrayPoint(3, i)) 'конец длинного штриха (подошва)
                    If PointPod1.X = 0 And PointPod1.Y = 0 Then
                        GoTo line4
                    End If
                    'строим длинный штрих
                    Dim GlobStart As Double = m / 1000
                    Dim acLine As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                    acLine.AddVertexAt(0, PointBrv1, 0, GlobStart, 0)
                    acLine.AddVertexAt(1, PointPod1, 0, 0, 0)
                    acLine.Layer = nameLayer
                    acLine.ColorIndex = nameColor
                    acLine.LineWeight = WidthSh
                    acLine.SetDatabaseDefaults()
                    acBlkTblRec.AppendEntity(acLine)
                    acTrans.AddNewlyCreatedDBObject(acLine, True)
                    'добавляем словарь в примитив
                    Dim extId As ObjectId = acLine.ExtensionDictionary
                    If extId.IsNull = True Then
                        acLine.UpgradeOpen()
                        acLine.CreateExtensionDictionary()
                        extId = acLine.ExtensionDictionary
                    End If
                    Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                    dbExt.UpgradeOpen()
                    Dim xRec As Xrecord = New Xrecord()
                    Dim rb As ResultBuffer = New ResultBuffer()
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Длинный штрих")) 'тип штриха
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                    rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                    xRec.Data = rb
                    dbExt.SetAt("SLOPEUSER", xRec)
                    acTrans.AddNewlyCreatedDBObject(xRec, True)
line4:
                Next
            ElseIf TipOtkos Like "Стенка подпорная деревянная" Then
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1)
                    Dim PointBrv1 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i)) 'начало длинного штриха (бровка)
                    Dim PointPod1 As Point2d = New Point2d(ArrayPoint(2, i), ArrayPoint(3, i)) 'конец длинного штриха (подошва)
                    If PointPod1.X = 0 And PointPod1.Y = 0 Then
                        GoTo line5
                    End If
                    'строим длинный штрих
                    Dim GlobStart As Double = m / 1000
                    GlobStart = GlobStart / 2
                    'находим две точки дополнително
                    Dim angle As Double = PointBrv1.GetVectorTo(PointPod1).Angle
                    Dim dx1 As Double = GlobStart * Math.Cos(angle + Math.PI / 2)
                    Dim dy1 As Double = GlobStart * Math.Sin(angle + Math.PI / 2)
                    Dim Pt1 As Point2d = New Point2d(PointBrv1.X + dx1, PointBrv1.Y + dy1)
                    Dim dx2 As Double = GlobStart * Math.Cos(angle - Math.PI / 2)
                    Dim dy2 As Double = GlobStart * Math.Sin(angle - Math.PI / 2)
                    Dim Pt2 As Point2d = New Point2d(PointBrv1.X - dx1, PointBrv1.Y - dy1)
                    Dim acLine As Autodesk.AutoCAD.DatabaseServices.Polyline = New Autodesk.AutoCAD.DatabaseServices.Polyline
                    acLine.AddVertexAt(0, Pt1, 0, 0, 0)
                    acLine.AddVertexAt(1, PointBrv1, 0, 0, 0)
                    acLine.AddVertexAt(2, Pt2, 0, 0, 0)
                    acLine.AddVertexAt(3, PointPod1, 0, 0, 0)
                    acLine.Closed = True
                    acLine.Layer = nameLayer
                    acLine.ColorIndex = nameColor
                    acLine.LineWeight = WidthSh
                    acLine.SetDatabaseDefaults()
                    acBlkTblRec.AppendEntity(acLine)
                    acTrans.AddNewlyCreatedDBObject(acLine, True)
                    'добавляем словарь в примитив
                    Dim extId As ObjectId = acLine.ExtensionDictionary
                    If extId.IsNull = True Then
                        acLine.UpgradeOpen()
                        acLine.CreateExtensionDictionary()
                        extId = acLine.ExtensionDictionary
                    End If
                    Dim dbExt As DBDictionary = acTrans.GetObject(extId, OpenMode.ForRead)
                    dbExt.UpgradeOpen()
                    Dim xRec As Xrecord = New Xrecord()
                    Dim rb As ResultBuffer = New ResultBuffer()
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Otkos")) 'имя начальной группы
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, TipOtkos)) 'тип откоса
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "Длинный штрих")) 'тип штриха
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, guidOtkos)) 'GUID
                    rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, HandleBrv)) 'handle бровки
                    rb.Add(New TypedValue(DxfCode.ExtendedDataReal, m)) 'масштаб
                    xRec.Data = rb
                    dbExt.SetAt("SLOPEUSER", xRec)
                    acTrans.AddNewlyCreatedDBObject(xRec, True)
line5:
                Next
            End If
            acTrans.Commit()
        End Using
        FuncOtkosDraw = True
    End Function
    '===============================================================================================
    'рисование откосов (перпендикуляр от бровки)
    Public Sub AutoDrawSlope(Optional ByVal lsh As Double = 1, Optional ByVal nameLn As String = "Откос неукрепленный", Optional ByVal nameLayer As String = "", Optional ByVal nameColor As String = "7", Optional ByVal DistSh As Double = 0.5, Optional ByVal rnd As Integer = 0, Optional M As Double = 500, Optional ByVal ShBool As Boolean = False, Optional ByVal WPlineSh As Double = 20)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        ' Create and display the value of two GUIDs.
        Dim g As Guid
        g = Guid.NewGuid()
        'линии для бровки и подошвы
        Dim PlineBrv As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
        Dim PlinePod As Autodesk.AutoCAD.DatabaseServices.Polyline = Nothing
        Dim PlineBrv2d As Polyline2d = Nothing
        Dim PlinePod2d As Polyline2d = Nothing
        Dim CurveBrv As Curve = Nothing
        Dim CurvePod As Curve = Nothing
        'Handle линий
        Dim PlineBrvHandle As String = Nothing
        Dim PlinePodHandle As String = Nothing
        Dim DistBrv As Double = 0
        Dim DistPod As Double = 0
        Dim PointPline As Point3d = New Point3d
        Dim PointPline2 As Point3d = New Point3d
        Dim ArrayPoint As Double(,) = Nothing
        Dim nVert As Integer = 1
        Dim Count As Integer = 0
        Dim NumVertPline As Integer = 1
        Dim ElevBrv As Double = 0
        Dim ElevPod As Double = 0
        Dim StartPointBrv As Double = 0
        Dim EndPointBrv As Double = 0
        Dim docLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем бровку откоса
            Dim SelectBrovka As PromptEntityResult = ed.GetEntity("Выберите бровку откоса: ")
            If SelectBrovka.Status = PromptStatus.Cancel Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectBrovka.ObjectId, OpenMode.ForWrite)
            Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
            If obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline" Then
                PlineBrv = obj
                ElevBrv = PlineBrv.Elevation 'запоминаем уровень полилинии, меняем на 0
                PlineBrv.Elevation = 0
                DistBrv = PlineBrv.Length 'длина полилинии
                CurveBrv = PlineBrv 'конвертируем в дугу
            ElseIf obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline2d" Then
                PlineBrv2d = obj
                ElevBrv = PlineBrv2d.Elevation 'запоминаем уровень полилинии, меняем на 0
                PlineBrv2d.Elevation = 0
                DistBrv = PlineBrv2d.Length 'длина полилинии
                CurveBrv = PlineBrv2d 'конвертируем в дугу
            Else
                MsgBox("Выбранный объект не является полилинией или 2д полилинией. Мокрос завершен")
                Exit Sub
            End If
            PlineBrvHandle = obj.Handle.ToString 'Handle бровки
            'выбираем подошву откоса
            Dim SelectPodoshva As PromptEntityResult = ed.GetEntity("Выберите подошву откоса: ")
            If SelectPodoshva.Status = PromptStatus.Cancel Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            Dim obj2 As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectPodoshva.ObjectId, OpenMode.ForWrite)
            Dim bt2 As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
            If obj2.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline" Then
                PlinePod = obj2
                ElevPod = PlinePod.Elevation 'запоминаем уровень полилинии, меняем на 0
                PlinePod.Elevation = 0
                DistPod = PlinePod.Length
                CurvePod = PlinePod
            ElseIf obj2.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline2d" Then
                PlinePod2d = obj2
                ElevPod = PlinePod2d.Elevation 'запоминаем уровень полилинии, меняем на 0
                PlinePod2d.Elevation = 0
                DistPod = PlinePod2d.Length
                CurvePod = PlinePod2d
            Else
                MsgBox("Выбранный объект не является полилинией или 2д полилинией. Мокрос завершен")
                Exit Sub
            End If
            PlinePodHandle = obj2.Handle.ToString 'Handle подошвы
            'запрос на обработку всей выбранной полилинии или только ее части
            Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
            pIntOpts.Message = vbCrLf & "Отрисовать откос? "
            pIntOpts.Keywords.Add("Да")
            pIntOpts.Keywords.Add("Параметры")
            pIntOpts.Keywords.Default = "Да"
            pIntOpts.AllowNone = True
            Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
            If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            If pIntRes.StringResult Like "Да" Then
                StartPointBrv = 0
                EndPointBrv = DistBrv
            Else
                'начальный штрих
                Try
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.Message = vbLf & "Укажите точку начала рисования откоса на бровке: "
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.None Then
                        StartPointBrv = 0
                    Else
                        Dim ptStart As Point3d = CurveBrv.GetClosestPointTo(pPtRes.Value, True)
                        StartPointBrv = CurveBrv.GetDistAtPoint(ptStart)
                    End If
                Catch ex As Exception
                    StartPointBrv = 0
                End Try
                'конечный штрих
                Try
                    Dim pPtRes2 As PromptPointResult
                    Dim pPtOpts2 As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts2.Message = vbLf & "Укажите точку конца отрисовки откоса на бровке: "
                    pPtRes2 = acDoc.Editor.GetPoint(pPtOpts2)
                    If pPtRes2.Status = PromptStatus.None Then
                        EndPointBrv = DistBrv
                    Else
                        Dim ptStart2 As Point3d = CurveBrv.GetClosestPointTo(pPtRes2.Value, True)
                        EndPointBrv = CurveBrv.GetDistAtPoint(ptStart2)
                    End If
                    'формирование массива точек для бровки
                Catch ex As Exception
                    EndPointBrv = DistBrv
                End Try
                'условие на случай если конец рисования откоса меньше его начала
                If StartPointBrv > EndPointBrv Then
                    Dim temp As Double = EndPointBrv
                    EndPointBrv = StartPointBrv
                    StartPointBrv = temp
                End If
            End If
            '=======================================================================================
            'расчет точек
            If nameLn Like "Откос неукрепленный" Or nameLn Like "Откос укрепленный" Then
                DistSh = DistSh * 2
            End If
            NumVertPline = PlinePod.NumberOfVertices
            For i As Double = StartPointBrv To Int(EndPointBrv) Step DistSh
                PointPline = CurveBrv.GetPointAtDist(i)
                If i > 0 Then
                    PointPline2 = CurveBrv.GetPointAtDist(i - DistSh)
                Else
                    PointPline2 = CurveBrv.GetPointAtDist(i + DistSh)
                End If
                Dim Pt1 As Point2d = New Point2d(PointPline.X, PointPline.Y)
                Dim Pt2 As Point2d = New Point2d(PointPline2.X, PointPline2.Y)
                Dim Angle As Double = Pt1.GetVectorTo(Pt2).Angle
                Angle = Angle - Math.PI / 2
                If Angle < 0 Then Angle = Angle + Math.PI * 2
                Dim Pt2X As Double = lsh * Math.Cos(Angle) + PointPline.X
                Dim Pt2Y As Double = lsh * Math.Sin(Angle) + PointPline.Y
                Dim point2 As Point3d = New Point3d(Pt2X, Pt2Y, 0)
                Dim Line1 As Line = New Line(PointPline, point2)
                Dim intersectPoint As Point3dCollection = New Point3dCollection
                Line1.IntersectWith(CurvePod, Intersect.ExtendThis, New Plane(), intersectPoint, IntPtr.Zero, IntPtr.Zero)
                If intersectPoint.Count > 0 Then
                    Dim Dist As Double = 1000000000000 'бесконечность
                    Dim Flag As Boolean = False
                    For Each pt As Point3d In intersectPoint
                        Dim lenghtPt As Double = Math.Sqrt((pt.X - PointPline.X) ^ 2 + (pt.Y - PointPline.Y) ^ 2)
                        If lenghtPt < Dist Then
                            ReDim Preserve ArrayPoint(3, Count)
                            ArrayPoint(0, Count) = PointPline.X
                            ArrayPoint(1, Count) = PointPline.Y
                            ArrayPoint(2, Count) = pt.X
                            ArrayPoint(3, Count) = pt.Y
                            Dist = lenghtPt
                            Flag = True
                        End If
                    Next
                    If Flag = True Then Count += 1
                Else
                    ReDim Preserve ArrayPoint(3, Count)
                    ArrayPoint(0, Count) = PointPline.X
                    ArrayPoint(1, Count) = PointPline.Y
                    Count += 1
                End If
            Next i
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'проверяем массив на ошибки и исправляем его
            If IsNothing(ArrayPoint) = False Then
                Dim StartOtkosSh As Boolean = False 'флаг (означает что построение откоса начато)
                Dim ArrayErrorPoint As Double(,) = Nothing
                Dim CountArrayErrorPoint As Integer = 0
                For i As Integer = 0 To ArrayPoint.GetUpperBound(1) - 1
                    'берем первый штрих
                    Dim PtBrv1 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i)) 'точка на бровке 
                    Dim PtPod1 As Point2d = New Point2d(ArrayPoint(2, i), ArrayPoint(3, i)) 'точка на подошве
                    Dim PtBrv2 As Point2d = New Point2d(ArrayPoint(0, i + 1), ArrayPoint(1, i + 1)) 'точка на бровке 
                    Dim PtPod2 As Point2d = New Point2d(ArrayPoint(2, i + 1), ArrayPoint(3, i + 1)) 'точка на подошве
                    'если это начало откоса и на подошве нет точек то игнорируем)
                    If PtPod1.X = 0 And PtPod1.Y = 0 And StartOtkosSh = False Then
                        GoTo line23
                        'если на подошве точки не найдены, но это середина или конец бровки то считаем такие точки ошибочными и записываем их в массив
                    End If
                    'следующий штрих не найден
                    If PtPod1.X <> 0 And PtPod1.Y <> 0 And PtPod2.X = 0 And PtPod2.Y = 0 Then
                        'смотрим если массив ошибок не пустой тогда сравниваем эту линию с последней линией в массиве
                        If IsArray(ArrayErrorPoint) = True Then
                            Dim i2 As Integer = ArrayErrorPoint.GetUpperBound(1)
                            Dim ptArrayBrv As Point2d = New Point2d(ArrayErrorPoint(0, i2), ArrayErrorPoint(1, i2))
                            Dim ptArrayPod As Point2d = New Point2d(ArrayErrorPoint(2, i2), ArrayErrorPoint(3, i2))
                            Dim PtIntersect1 As Point2d = Nothing
                            Dim boolIntersect As Boolean = MathFunction.FuncIntersectPolyline(PtBrv1, PtPod1, ptArrayBrv, ptArrayPod, PtIntersect1)
                            If boolIntersect = True Then
                                If PtIntersect1.X <> 0 And PtIntersect1.Y <> 0 Then
                                    ReDim Preserve ArrayErrorPoint(3, CountArrayErrorPoint)
                                    ArrayErrorPoint(0, CountArrayErrorPoint) = PtBrv1.X
                                    ArrayErrorPoint(1, CountArrayErrorPoint) = PtBrv1.Y
                                    ArrayErrorPoint(2, CountArrayErrorPoint) = PtPod1.X
                                    ArrayErrorPoint(3, CountArrayErrorPoint) = PtPod1.Y
                                    CountArrayErrorPoint += 1
                                    StartOtkosSh = True
                                    GoTo line23
                                End If
                            End If
                        End If
                        StartOtkosSh = True
                        GoTo line23
                    End If
                    'если штрих не найден но построение начато
                    If PtPod1.X = 0 And PtPod1.Y = 0 And StartOtkosSh = True Then
                        ReDim Preserve ArrayErrorPoint(3, CountArrayErrorPoint)
                        ArrayErrorPoint(0, CountArrayErrorPoint) = PtBrv1.X
                        ArrayErrorPoint(1, CountArrayErrorPoint) = PtBrv1.Y
                        ArrayErrorPoint(2, CountArrayErrorPoint) = PtPod1.X
                        ArrayErrorPoint(3, CountArrayErrorPoint) = PtPod1.Y
                        CountArrayErrorPoint += 1
                        GoTo line23
                        'если на подошве точки не найдены, но это середина или конец бровки то считаем такие точки ошибочными и записываем их в массив
                    End If
                    'проверяем на пересечение
                    Dim PtIntersect As Point2d = Nothing
                    Dim boolIntersct1 As Boolean = MathFunction.FuncIntersectPolyline(PtBrv1, PtPod1, PtBrv2, PtPod2, PtIntersect)
                    If PtIntersect.X = 0 And PtIntersect.Y = 0 Then 'если пересечения нет тогда
                        If IsNothing(ArrayErrorPoint) = False Then
                            'проверяем этот штрих на пересечение со всеми элементами в массиве
                            Dim boolPer As Boolean = False
                            For i2 As Integer = 0 To ArrayErrorPoint.GetUpperBound(1)
                                Dim ptArrayBrv As Point2d = New Point2d(ArrayErrorPoint(0, i2), ArrayErrorPoint(1, i2))
                                Dim ptArrayPod As Point2d = New Point2d(ArrayErrorPoint(2, i2), ArrayErrorPoint(3, i2))
                                If ptArrayPod.X <> 0 And ptArrayPod.Y <> 0 Then
                                    Dim PtIntersect1 As Point2d = Nothing
                                    Dim boolIntr As Boolean = MathFunction.FuncIntersectPolyline(PtBrv1, PtPod1, ptArrayBrv, ptArrayPod, PtIntersect1)
                                    If PtIntersect1.X <> 0 And PtIntersect1.Y <> 0 Then
                                        ReDim Preserve ArrayErrorPoint(3, CountArrayErrorPoint)
                                        ArrayErrorPoint(0, CountArrayErrorPoint) = PtBrv1.X
                                        ArrayErrorPoint(1, CountArrayErrorPoint) = PtBrv1.Y
                                        ArrayErrorPoint(2, CountArrayErrorPoint) = PtPod1.X
                                        ArrayErrorPoint(3, CountArrayErrorPoint) = PtPod1.Y
                                        CountArrayErrorPoint += 1
                                        StartOtkosSh = True
                                        GoTo line23
                                    End If
                                End If
                            Next
                        End If
                        'если пересечений нет тогда
                        If IsNothing(ArrayErrorPoint) = False Then
                            Dim IndStart As Integer = i - ArrayErrorPoint.GetUpperBound(1) - 2
                            Dim IndEnd As Integer = i
                            Dim TempStartShPrev As Point3d = New Point3d(ArrayPoint(2, IndStart), ArrayPoint(3, IndStart), 0)
                            Dim TempEndPointPod As Point3d = New Point3d(ArrayPoint(2, IndEnd), ArrayPoint(3, IndEnd), 0)
                            Dim StartShPrev As Point3d = CurvePod.GetClosestPointTo(TempStartShPrev, True)
                            Dim EndPointPod As Point3d = CurvePod.GetClosestPointTo(TempEndPointPod, True)
                            Dim DistStart As Double = CurvePod.GetDistAtPoint(StartShPrev)
                            Dim DistEnd As Double = CurvePod.GetDistAtPoint(EndPointPod)
                            Dim DistStartEndPod As Double = Math.Abs(DistEnd - DistStart)
                            Dim DeltaDist As Double = DistStartEndPod / (ArrayErrorPoint.GetUpperBound(1) + 1)
                            'вытаскиваем из первого массива предыдущий штрих на котором пересечения небыло
                            For k As Integer = 0 To ArrayErrorPoint.GetUpperBound(1)  'записываем в первый массив вычисленные точки по подошве
                                DistStart = DistStart + DeltaDist
                                Dim PointPodNext As Point3d = CurvePod.GetPointAtDist(DistStart)
                                ArrayPoint(2, k + 1 + IndStart) = PointPodNext.X
                                ArrayPoint(3, k + 1 + IndStart) = PointPodNext.Y
                            Next
                            Erase ArrayErrorPoint
                            CountArrayErrorPoint = 0
                        End If
                        StartOtkosSh = True
                        GoTo line23
                    Else
                        ReDim Preserve ArrayErrorPoint(3, CountArrayErrorPoint)
                        ArrayErrorPoint(0, CountArrayErrorPoint) = PtBrv1.X
                        ArrayErrorPoint(1, CountArrayErrorPoint) = PtBrv1.Y
                        ArrayErrorPoint(2, CountArrayErrorPoint) = PtPod1.X
                        ArrayErrorPoint(3, CountArrayErrorPoint) = PtPod1.Y
                        CountArrayErrorPoint += 1
                        StartOtkosSh = True
                    End If
line23:
                Next
            End If
            'рисуем откос
            Dim boolOtkos As Boolean = FuncOtkosDraw(ArrayPoint, CurveBrv, CurvePod, nameLn, lsh, nameLayer, nameColor, rnd, CurveBrv.Handle.Value.ToString, M, ShBool)
            docLock.Dispose()
            acTrans.Commit()
        End Using
    End Sub
    'рисование откоса вручную
    Public Sub ManuallyDrawSlope(Optional ByVal lsh As Double = 0.5, Optional ByVal nameLn As String = "Откос неукрепленный", Optional ByVal nameLayer As String = "0", Optional ByVal nameColor As String = "7", Optional ByVal DistSh As Double = 1)
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim PlineBrv As Autodesk.AutoCAD.DatabaseServices.Polyline
        Dim Pline2dBrv As Polyline2d
        Dim CurveBrv As Curve
        Dim DistBrv As Double = 0
        Dim PointPline As Point3d = New Point3d
        Dim StartPointBrv As Double = 0
        Dim ArrayPointBrv As Double(,) = Nothing
        Dim Count As Integer = 0
        Dim Angle1 As Double = 0
        'переменные для первого штриха
        Dim ptStart1 As Point3d = Nothing
        Dim ptEnd1 As Point3d = Nothing
        Dim LenSh1 As Double = Nothing 'длина линии 1
        'переменные для второго штриха
        Dim ptStart2 As Point3d = Nothing
        Dim ptEnd2 As Point3d = Nothing
        Dim LenSh2 As Double = Nothing 'длина линии 2
        'запросы
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        Dim acLine As Line = New Line()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim docLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
            Dim SelectBrovka As PromptEntityResult = ed.GetEntity("Выберите бровку откоса: ")
            If SelectBrovka.Status = PromptStatus.Cancel Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = acTrans.GetObject(SelectBrovka.ObjectId, OpenMode.ForWrite)
            Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
            If obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline" Then
                PlineBrv = obj
                DistBrv = PlineBrv.Length 'длина полилинии
                CurveBrv = PlineBrv 'конвертируем в дугу
            ElseIf obj.GetType().ToString = "Autodesk.AutoCAD.DatabaseServices.Polyline2d" Then
                Pline2dBrv = obj
                DistBrv = Pline2dBrv.Length 'длина полилинии
                CurveBrv = Pline2dBrv 'конвертируем в дугу
            Else
                MsgBox("Вабранный объект не является полилинией или 2d полилинией. Мокрос завершен")
                Exit Sub
            End If
            'начальный штрих

            pPtOpts.Message = vbLf & "Укажите точку начала откоса на бровке откоса: "
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            ptStart1 = pPtRes.Value
            StartPointBrv = CurveBrv.GetDistAtPoint(ptStart1)
            If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
            pPtOpts.Message = vbLf & "Укажите точку начала откоса на подошве откоса: "
            pPtOpts.UseBasePoint = True
            pPtOpts.BasePoint = ptStart1
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            ptEnd1 = pPtRes.Value
            If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
            acLine = New Line(ptStart1, ptEnd1)
            acLine.Layer = nameLayer
            acLine.ColorIndex = nameColor
            acLine.LineWeight = LineWeight.LineWeight020
            acLine.SetDatabaseDefaults()
            acBlkTblRec.AppendEntity(acLine)
            acTrans.AddNewlyCreatedDBObject(acLine, True)
            Angle1 = acLine.Angle 'запоминаем угол 1 линии
            LenSh1 = acLine.Length
            acTrans.Commit()
            docLock.Dispose()
        End Using
        If nameLn Like "Обрыв земляной" Then
            'организация неопределенного цикла
            While ptEnd1 <> Nothing
                Try
                    ptEnd2 = Nothing
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        Dim acBlkTblRec As BlockTableRecord
                        acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Dim docLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                        StartPointBrv = StartPointBrv + DistSh
                        ptStart2 = CurveBrv.GetPointAtDist(StartPointBrv)
                        pPtOpts.Message = vbLf & "Укажите следующую точку на подошве откоса: "
                        pPtOpts.UseBasePoint = True
                        pPtOpts.BasePoint = ptStart2
                        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                        ptEnd2 = pPtRes.Value
                        If pPtRes.Status = PromptStatus.Cancel Then Exit While
                        acLine = New Line(ptStart2, ptEnd2)
                        acLine.Layer = nameLayer
                        acLine.ColorIndex = nameColor
                        acLine.LineWeight = LineWeight.LineWeight020
                        acLine.SetDatabaseDefaults()
                        acBlkTblRec.AppendEntity(acLine)
                        acTrans.AddNewlyCreatedDBObject(acLine, True)
                        ptStart1 = ptStart2
                        docLock.Dispose()
                        acTrans.Commit()
                    End Using
                Catch
                    Exit While
                End Try
            End While
        Else
            'организация неопределенного цикла
            While ptEnd1 <> Nothing
                Try
                    ptEnd2 = Nothing
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable
                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                        Dim acBlkTblRec As BlockTableRecord
                        acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Dim docLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
                        StartPointBrv = StartPointBrv + DistSh * 2
                        ptStart2 = CurveBrv.GetPointAtDist(StartPointBrv)
                        pPtOpts.Message = vbLf & "Укажите следующую точку на подошве откоса: "
                        pPtOpts.UseBasePoint = True
                        pPtOpts.BasePoint = ptStart2
                        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                        ptEnd2 = pPtRes.Value
                        If pPtRes.Status = PromptStatus.Cancel Then Exit While
                        acLine = New Line(ptStart2, ptEnd2)
                        acLine.Layer = nameLayer
                        acLine.ColorIndex = nameColor
                        acLine.LineWeight = LineWeight.LineWeight020
                        acLine.SetDatabaseDefaults()
                        acBlkTblRec.AppendEntity(acLine)
                        acTrans.AddNewlyCreatedDBObject(acLine, True)
                        Dim Angle2 As Double = acLine.Angle 'запоминаем угол 2 линии
                        LenSh2 = acLine.Length
                        'блок для короткого штриха
                        'вычисляем среднее из углов для определения дирекционного направления короткого штриха
                        Dim AngleSr As Double = (Angle1 + Angle2) / 2
                        Dim DirectA As Double = FuncSurvey.FuncConvertMathAngleToDirectionAngle(AngleSr)
                        Angle1 = Angle2
                        'проверка длины короткого штриха
                        Dim LenSh As Double = (LenSh1 + LenSh2) / 2
                        If LenSh < lsh Then
                            Dim dy As Double = LenSh * Math.Cos(DirectA)
                            Dim dx As Double = LenSh * Math.Sin(DirectA)
                            PointPline = CurveBrv.GetPointAtDist(StartPointBrv - DistSh)
                            Dim ptEndSh As Point3d = New Point3d(PointPline.X + dx, PointPline.Y + dy, 0) 'расчет конечной точки
                            acLine = New Line(PointPline, ptEndSh)
                        Else
                            Dim dy As Double = lsh * Math.Cos(DirectA)
                            Dim dx As Double = lsh * Math.Sin(DirectA)
                            PointPline = CurveBrv.GetPointAtDist(StartPointBrv - DistSh)
                            Dim ptEndSh As Point3d = New Point3d(PointPline.X + dx, PointPline.Y + dy, 0) 'расчет конечной точки
                            acLine = New Line(PointPline, ptEndSh)
                        End If
                        acLine.Layer = nameLayer
                        acLine.ColorIndex = nameColor
                        acLine.LineWeight = LineWeight.LineWeight020
                        acLine.SetDatabaseDefaults()
                        acBlkTblRec.AppendEntity(acLine)
                        acTrans.AddNewlyCreatedDBObject(acLine, True)
                        'рисуем точку если выбран откос укрепленный
                        If nameLn Like "Откос укрепленный" And LenSh >= lsh * 1.5 Then
                            Dim dy As Double = 1.4 * lsh * Math.Cos(DirectA)
                            Dim dx As Double = 1.4 * lsh * Math.Sin(DirectA)
                            Dim PointStart As Point3d = New Point3d(PointPline.X + dx, PointPline.Y + dy, 0) 'расчет начальной точки
                            dy = 1.6 * lsh * Math.Cos(DirectA)
                            dx = 1.6 * lsh * Math.Sin(DirectA)
                            Dim PointEnd As Point3d = New Point3d(PointPline.X + dx, PointPline.Y + dy, 0) 'расчет начальной точки
                            acLine = New Line(PointStart, PointEnd)
                            acLine.Layer = nameLayer
                            acLine.ColorIndex = nameColor
                            acLine.LineWeight = LineWeight.LineWeight020
                            acLine.SetDatabaseDefaults()
                            acBlkTblRec.AppendEntity(acLine)
                            acTrans.AddNewlyCreatedDBObject(acLine, True)
                        End If
                        LenSh1 = LenSh2
                        ptStart1 = ptStart2
                        docLock.Dispose()
                        acTrans.Commit()
                    End Using
                Catch
                    Exit While
                End Try
            End While
        End If
    End Sub
    '===============================================================================================
    'рисование откосов
    <CommandMethod("DrawUserSlope")>
    Public Sub DrawUserSlope()
        'создаем класс с новой формой
        Dim otkosDraw As DrawUserSlopes = New DrawUserSlopes
        otkosDraw.ShowDialog()
        If otkosDraw.stbool = False Then Exit Sub
    End Sub
End Class
