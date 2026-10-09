Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Colors

Public Class FuncSurvey
    '==========================================================================
    'функция преобразут дирекционные направления в румбы (результат Румб ф формате град мин сек)
    Public Shared Function FuncDirectToRumbDDMMSS(ByVal DegGradus As Double, Optional ByVal roudZn As Integer = 0) As String
        DegGradus = Math.Abs(DegGradus)
        If DegGradus > 360 Then
            DegGradus = DegGradus - 360
        End If
        Dim Napr As String = ""
        Try
            If DegGradus = 0 Then
                Napr = "С: "
            ElseIf DegGradus = 90 Then
                Napr = "В: "
            ElseIf DegGradus = 180 Then
                Napr = "Ю: "
            ElseIf DegGradus = 270 Then
                Napr = "З: "
            Else
                Select Case DegGradus
                    Case 0 To 90
                        Napr = "СВ: "
                    Case 90 To 180
                        DegGradus = 180 - DegGradus
                        Napr = "ЮВ: "
                    Case 180 To 270
                        DegGradus = DegGradus - 180
                        Napr = "ЮЗ: "
                    Case 270 To 360
                        DegGradus = 360 - DegGradus
                        Napr = "СЗ: "
                    Case Else
                        FuncDirectToRumbDDMMSS = "0.0000"
                        Exit Function
                End Select
            End If
            '===================================================================================
            'начинаем преобразование в формат dd mm ss
            Dim Grad As Integer = Int(DegGradus)
            Dim Min As Integer = Int((DegGradus - Grad) * 60)
            Dim sec As Double = Math.Round((((DegGradus - Grad) * 60) - Min) * 60, roudZn)
            'начинаем проверку на случай если число минут и секунд равно 60
            If sec = 60 Then
                sec = 0
                Min = Min + 1
            End If
            If Min = 60 Then
                Min = 0
                Grad = Grad + 1
            End If
            Dim Min2 As String = String.Format("{0:00}", Min)
            Dim sec2 As String = sec
            If roudZn = 0 Then
                sec2 = String.Format("{0:00}", sec)
            ElseIf roudZn = 1 Then
                sec2 = String.Format("{0:00.0}", sec)
            ElseIf roudZn = 2 Then
                sec2 = String.Format("{0:00.00}", sec)
            ElseIf roudZn = 3 Then
                sec2 = String.Format("{0:00.000}", sec)
            ElseIf roudZn = 4 Then
                sec2 = String.Format("{0:00.0000}", sec)
            End If
            FuncDirectToRumbDDMMSS = Napr & Grad & Chr(176) & Min2 & "'" & sec2 & """"
        Catch
            FuncDirectToRumbDDMMSS = "0.0000"
            Exit Function
        End Try
    End Function
    '==========================================================================
    'функция переводит значения дирекционного направления из геодезической СК в математическую
    Public Shared Function FuncDirectionAngleToMathAngle(ByVal DirectionAngle As Double) As Double
        FuncDirectionAngleToMathAngle = 0
        Try
            Select Case DirectionAngle
                Case 0 To Math.PI / 2
                    FuncDirectionAngleToMathAngle = Math.PI / 2 - DirectionAngle
                Case Math.PI / 2 To Math.PI
                    FuncDirectionAngleToMathAngle = 2 * Math.PI - DirectionAngle + Math.PI / 2
                Case Math.PI To (3 * Math.PI) / 2
                    FuncDirectionAngleToMathAngle = (Math.PI * 2 - DirectionAngle) + Math.PI / 2
                Case (3 * Math.PI) / 2 To Math.PI * 2
                    FuncDirectionAngleToMathAngle = 2 * Math.PI - DirectionAngle + Math.PI / 2
                Case Else
                    FuncDirectionAngleToMathAngle = 0
            End Select
        Catch
            FuncDirectionAngleToMathAngle = 0
            Exit Function
        End Try
    End Function
    '==========================================================================
    'функция переводит значения дирекционного направления из математической СК в геодезическую
    Public Shared Function FuncConvertMathAngleToDirectionAngle(ByVal DegMathAngle As Double) As Double
        Try
            Select Case DegMathAngle
                Case 0 To Math.PI / 2
                    FuncConvertMathAngleToDirectionAngle = Math.PI / 2 - DegMathAngle
                Case Math.PI / 2 To Math.PI
                    FuncConvertMathAngleToDirectionAngle = 2 * Math.PI - DegMathAngle + Math.PI / 2
                Case Math.PI To 1.5 * Math.PI
                    FuncConvertMathAngleToDirectionAngle = 2 * Math.PI - DegMathAngle + Math.PI / 2
                Case 1.5 * Math.PI To Math.PI * 2
                    FuncConvertMathAngleToDirectionAngle = 2 * Math.PI - DegMathAngle + Math.PI / 2
            End Select
        Catch
            FuncConvertMathAngleToDirectionAngle = 0
            Exit Function
        End Try
    End Function
    '================================================================================================
    'функция для поиска CoGo точки с определенными координатами
    Public Shared Function FuncFindCogoPointByCoordinates(ByVal PointCoord As Point3d, ByVal NumPoint As Integer(), Optional ByVal dopusk As Double = 0, Optional PTGroup As PointGroup = Nothing) As CogoPoint
        FuncFindCogoPointByCoordinates = Nothing
        If IsArray(NumPoint) = False Then
            Exit Function
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim Lenght As Double = 10000000
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Try
                For i As Integer = 0 To NumPoint.GetUpperBound(0)
                    Dim IDCoGoPoint As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(NumPoint(i))
                    If IsNothing(IDCoGoPoint) = True Then Continue For
                    If IDCoGoPoint.IsNull = True Then Continue For
                    Dim CoGoPoint As CogoPoint = IDCoGoPoint.GetObject(OpenMode.ForRead)
                    If IsNothing(PTGroup) = False Then
                        If PTGroup.ContainsPoint(CoGoPoint.PointNumber) = False Then
                            Continue For
                        End If
                    End If
                    Dim t1 As Point2d = New Point2d(PointCoord.X, PointCoord.Y)
                    Dim t2 As Point2d = New Point2d(CoGoPoint.Easting, CoGoPoint.Northing)
                    Dim Dist As Double = t1.GetVectorTo(t2).Length
                    Dist = Math.Round(Dist, 3)
                    If Dist = 0 Then
                        FuncFindCogoPointByCoordinates = CoGoPoint
                        Exit Function
                    ElseIf Dist < Lenght Then
                        Lenght = Dist
                        FuncFindCogoPointByCoordinates = CoGoPoint
                    End If
                Next
                If Lenght > dopusk Then
                    FuncFindCogoPointByCoordinates = Nothing
                End If
            Catch
                FuncFindCogoPointByCoordinates = Nothing
            End Try
            tr.Commit()
        End Using
    End Function
    '=====================================================================================================
    'функция для поиска ближайшей COGO точки к выбранной COGO точке по расстоянию
    Public Shared Function FuncFindCogoPointByDistance(ByVal Numb As Integer, Optional ByVal dopusk As Double = 0.2) As CogoPoint
        FuncFindCogoPointByDistance = Nothing
        Dim Lenght As Double = 10000000
        Dim PointId As ObjectId = Nothing
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Try
                'находим координаты исходной точки
                Dim PointIdIsh As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(Numb)
                Dim PointCOGOIsh As CogoPoint = PointIdIsh.GetObject(OpenMode.ForRead)
                Dim t1 As Point2d = New Point2d(PointCOGOIsh.Easting, PointCOGOIsh.Northing)
                For Each PointId In CivilDoc.CogoPoints
                    Dim PointCOGO1 As CogoPoint = PointId.GetObject(OpenMode.ForRead)
                    If Not (PointCOGO1.PointNumber = Numb) Then
                        Dim X1 As Double = PointCOGO1.Easting
                        Dim Y1 As Double = PointCOGO1.Northing
                        Dim t2 As Point2d = New Point2d(X1, Y1)
                        Dim Dist As Double = t1.GetVectorTo(t2).Length
                        Dist = Math.Round(Dist, 3)
                        If Dist < Lenght Then
                            Lenght = Dist
                            FuncFindCogoPointByDistance = PointCOGO1
                        End If
                    End If
                Next
                If Lenght > dopusk Then
                    FuncFindCogoPointByDistance = Nothing
                End If
            Catch
                FuncFindCogoPointByDistance = Nothing
            End Try
        End Using
    End Function
    '-==================================================================================
    'функция вычисляет поворот маркера COGO точки по точкам полилинии (для функции бордюрный камень)
    Public Shared Function FuncCalcRotateMarkerCoGoPoint(ByRef ArrayPoint(,) As Double, ByVal x As Double, ByVal y As Double) As Double
        FuncCalcRotateMarkerCoGoPoint = 0 'значение по умолчанию
        Dim dist As Double = 90000000
        Dim NumbVertex As Integer = 0
        Dim P1 As Point2d = New Point2d(x, y) 'положение точки низа бордюрного камня
        'цикл для опроеделения в массиве индекса для этой точки
        For i As Integer = 0 To ArrayPoint.GetUpperBound(1) - 1
            Dim P2 As Point2d = New Point2d(ArrayPoint(0, i), ArrayPoint(1, i))
            Dim dist1 As Double = P2.GetVectorTo(P1).Length 'вычисляем расстояние
            If dist1 = 0 Then 'если расстояние равно 0 это значит что точка найдена выходим из цикла
                NumbVertex = i
                dist = 0
                Exit For
            Else
                If dist > dist1 Then 'иначе держим в памяти расстояние для нахождения наименьшего
                    dist = dist1
                    NumbVertex = i
                End If
            End If
        Next i
        Dim t1 As Point2d = New Point2d() 'предыдущая точка
        Dim t2 As Point2d = New Point2d() 'центральная точка
        Dim t3 As Point2d = New Point2d() 'последующая точка
        If dist < 0.1 Then
            Dim angle As Double = 0
            If NumbVertex = 0 Then
                t1 = New Point2d(ArrayPoint(0, NumbVertex), ArrayPoint(1, NumbVertex))
                t2 = New Point2d(ArrayPoint(0, NumbVertex + 1), ArrayPoint(1, NumbVertex + 1))
                angle = t2.GetVectorTo(t1).Angle
                If angle > Math.PI / 2 And angle <= Math.PI Then
                    angle = angle + Math.PI
                ElseIf angle > Math.PI And angle <= (3 * Math.PI) / 2 Then
                    angle = angle - Math.PI
                End If
            ElseIf NumbVertex = ArrayPoint.GetUpperBound(1) Then
                t1 = New Point2d(ArrayPoint(0, NumbVertex), ArrayPoint(1, NumbVertex))
                t2 = New Point2d(ArrayPoint(0, NumbVertex - 1), ArrayPoint(1, NumbVertex - 1))
                angle = t1.GetVectorTo(t2).Angle
                If angle > Math.PI / 2 And angle <= Math.PI Then
                    angle = angle + Math.PI
                ElseIf angle > Math.PI And angle <= (3 * Math.PI) / 2 Then
                    angle = angle - Math.PI
                End If
            Else
                t1 = New Point2d(ArrayPoint(0, NumbVertex - 1), ArrayPoint(1, NumbVertex - 1))
                t2 = New Point2d(ArrayPoint(0, NumbVertex), ArrayPoint(1, NumbVertex))
                t3 = New Point2d(ArrayPoint(0, NumbVertex + 1), ArrayPoint(1, NumbVertex + 1))
                angle = t2.GetVectorTo(t1).Angle
                If angle > Math.PI / 2 And angle <= Math.PI Then
                    angle = t2.GetVectorTo(t3).Angle
                ElseIf angle > Math.PI And angle <= (3 * Math.PI) / 2 Then
                    angle = t2.GetVectorTo(t3).Angle
                End If
            End If
            FuncCalcRotateMarkerCoGoPoint = angle
        Else
            FuncCalcRotateMarkerCoGoPoint = 0
        End If
    End Function
    '=====================================================================================================
    'функция возвращает наибольший номер COGO точки
    Public Shared Function FuncFindMaxNumberCoGoPoint() As Integer
        FuncFindMaxNumberCoGoPoint = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Try
                For Each id As ObjectId In CivilDoc.CogoPoints
                    Dim cogoPoint As CogoPoint = id.GetObject(OpenMode.ForRead)
                    If FuncFindMaxNumberCoGoPoint < cogoPoint.PointNumber Then
                        FuncFindMaxNumberCoGoPoint = cogoPoint.PointNumber
                    End If
                Next
            Catch
                FuncFindMaxNumberCoGoPoint = 0
            End Try
        End Using
    End Function
    '=====================================================================================================
    'функция проверяет есть ли COGO точка с таким номером
    Public Shared Function FuncReviseNumberCoGoPoint(ByVal Numb As Integer) As Boolean
        FuncReviseNumberCoGoPoint = False
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Try
            Dim PointIdIsh As ObjectId = CivilDoc.CogoPoints.GetPointByPointNumber(Numb)
            FuncReviseNumberCoGoPoint = True
        Catch
            FuncReviseNumberCoGoPoint = False
        End Try
    End Function
    '=====================================================================================================
    'функция возвращает заносит в массив номера точек COGO которые входят в габаритный контейнер
    Public Shared Function FuncFindNumberCoGoPointToBoundExtents(ByVal Ent As Autodesk.AutoCAD.DatabaseServices.Entity, ByRef ArrayPoint() As Integer, Optional Dopusk As Double = 10) As Integer
        FuncFindNumberCoGoPointToBoundExtents = 0
        Try
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim MinCoord As Point3d = Ent.GeometricExtents.MinPoint
            Dim MaxCoord As Point3d = Ent.GeometricExtents.MaxPoint
            Dim Count1 As Integer = 0
            Dim CountPoint As Integer = CivilDoc.CogoPoints.Count
            If CountPoint > 0 Then
                For Each PointCogoID As ObjectId In CivilDoc.CogoPoints
                    Dim PointCOGO1 As CogoPoint = PointCogoID.GetObject(OpenMode.ForRead)
                    If PointCOGO1.Northing >= MinCoord.Y - Dopusk And PointCOGO1.Easting >= MinCoord.X - Dopusk Then
                        If PointCOGO1.Northing <= MaxCoord.Y + Dopusk And PointCOGO1.Easting <= MaxCoord.X + Dopusk Then
                            ReDim Preserve ArrayPoint(Count1)
                            ArrayPoint(Count1) = PointCOGO1.PointNumber
                            Count1 += 1
                        End If
                    End If
                Next
            End If
            tr.Commit()
            FuncFindNumberCoGoPointToBoundExtents = Count1
        Catch
            FuncFindNumberCoGoPointToBoundExtents = 0
        End Try
    End Function
    '==================================================================================================
    'функция заносит в массив номера COGO точек которые попадают в габаритный контейнер объекта
    Public Shared Function FuncWriteArrayNumberCoGoPointToBoundExtentsObject(ByVal Ent As Autodesk.AutoCAD.DatabaseServices.Entity, ByRef ArrayPoint() As Integer, Optional Dopusk As Double = 10) As Integer
        FuncWriteArrayNumberCoGoPointToBoundExtentsObject = 0
        Try
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = acDoc.Database
            Dim ed As Editor = acDoc.Editor
            Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Dim MinCoord As Point3d = Ent.GeometricExtents.MinPoint
            Dim MaxCoord As Point3d = Ent.GeometricExtents.MaxPoint
            Dim Count1 As Integer = 0
            Dim CountPoint As Integer = CivilDoc.CogoPoints.Count
            If CountPoint > 0 Then
                For Each PointCogoID As ObjectId In CivilDoc.CogoPoints
                    Dim PointCOGO1 As CogoPoint = PointCogoID.GetObject(OpenMode.ForRead)
                    If PointCOGO1.Northing >= MinCoord.Y - Dopusk And PointCOGO1.Easting >= MinCoord.X - Dopusk Then
                        If PointCOGO1.Northing <= MaxCoord.Y + Dopusk And PointCOGO1.Easting <= MaxCoord.X + Dopusk Then
                            ReDim Preserve ArrayPoint(Count1)
                            ArrayPoint(Count1) = PointCOGO1.PointNumber
                            Count1 += 1
                        End If
                    End If
                Next
            End If
            tr.Commit()
            FuncWriteArrayNumberCoGoPointToBoundExtentsObject = Count1
        Catch
            FuncWriteArrayNumberCoGoPointToBoundExtentsObject = 0
        End Try
    End Function
    '=======================================================================================================
    'функция создает новую COGO точку
    Public Shared Function FuncCreateCoGoPoint(ByVal PtXYZ As Point3d, ByVal Desk As String, Optional number As Integer = 0, Optional ByVal elev As Double = -1, Optional name As String = "") As ObjectId
        FuncCreateCoGoPoint = Nothing
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim docBase As Database = docCol.MdiActiveDocument.Database
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument 'ссылка на активный чертеж
        Dim db As Database = acDoc.Database
        Dim NumbPoint As String() = Nothing
        Dim CountPoint As Integer = 0
        Try
            Dim PointCogoID As ObjectId = CivilDoc.CogoPoints.Add(PtXYZ, True) 'вставка одиночной точки
            Dim PointCOGO2 As CogoPoint = PointCogoID.GetObject(OpenMode.ForWrite)
            PointCOGO2.RawDescription = Desk
            If number > 0 Then
                If CivilDoc.CogoPoints.Contains(number) = False Then
                    PointCOGO2.PointNumber = number
                End If
            End If
            If elev <> -1 Then
                PointCOGO2.Elevation = elev
            End If
            If name.Trim.Length > 0 Then
                PointCOGO2.PointName = name
            End If
            FuncCreateCoGoPoint = PointCOGO2.ObjectId
            PointCOGO2.ApplyDescriptionKeys()
        Catch ex As System.Exception
            FuncCreateCoGoPoint = Nothing
        End Try
    End Function
    '===============================================================================================================================
    'функция для поиска конкретного значения пользовательких свойств для COGO точки
    Public Shared Function FuncReadUDPValue(ByVal PointCOGO As CogoPoint, ByVal NameField As String, ByVal NameClassUdp As String) As String
        FuncReadUDPValue = ""
        'читаем пользовательские свойства точки и заносим их в массив
        For Each IdUdp As UDP In CivilApplication.ActiveDocument.PointUDPs
            If IsNothing(IdUdp) = False Then
                Dim TempNameClassUdp As String = IdUdp.ClassificationName
                Dim NameFieldUdp As String = IdUdp.Name
                If TempNameClassUdp Like NameClassUdp Then
                    If NameFieldUdp Like NameField Then
                        If TypeOf IdUdp Is UDPDouble Then 'вещественное с двойной точностью
                            Try
                                Dim UDp As UDPDouble = TryCast(IdUdp, UDPDouble)
                                FuncReadUDPValue = PointCOGO.GetUDPValue(UDp).ToString
                                Exit Function
                            Catch ex As System.InvalidOperationException
                            End Try
                        ElseIf TypeOf IdUdp Is UDPBoolean Then 'false/true
                            Try
                                Dim UDp As UDPBoolean = TryCast(IdUdp, UDPBoolean)
                                FuncReadUDPValue = PointCOGO.GetUDPValue(UDp).ToString
                                Exit Function
                            Catch ex As System.InvalidOperationException
                            End Try
                        ElseIf TypeOf IdUdp Is UDPEnumeration Then '
                            Try
                                Dim UDp As UDPEnumeration = TryCast(IdUdp, UDPEnumeration)
                                FuncReadUDPValue = PointCOGO.GetUDPValue(UDp).ToString
                                Exit Function
                            Catch ex As System.InvalidOperationException
                            End Try
                        ElseIf TypeOf IdUdp Is UDPInteger Then 'целое
                            Try
                                Dim UDp As UDPInteger = TryCast(IdUdp, UDPInteger)
                                FuncReadUDPValue = PointCOGO.GetUDPValue(UDp).ToString
                                Exit Function
                            Catch ex As System.InvalidOperationException
                            End Try
                        ElseIf TypeOf IdUdp Is UDPString Then 'строка
                            Try
                                Dim UDp As UDPString = TryCast(IdUdp, UDPString)
                                FuncReadUDPValue = PointCOGO.GetUDPValue(UDp).ToString
                                Exit Function
                            Catch ex As System.InvalidOperationException
                            End Try
                        End If
                    End If
                End If
            End If
        Next
    End Function
    '===============================================================================================================================
    'функция перезаписывает определенное поле пользовательских данных для COGO точек
    Public Shared Function FuncWriteUDPValue(ByVal PointCOGO As CogoPoint, ByVal ClassNameUDP As String, ByVal NameFieldUDP As String, ByVal UDPval As String) As Boolean
        FuncWriteUDPValue = False
        For Each IdUdp As UDP In CivilApplication.ActiveDocument.PointUDPs
            If IsNothing(IdUdp) = False Then
                Dim NameUdp As String = IdUdp.Name
                Dim ClassUDP As String = IdUdp.ClassificationName
                If ClassUDP Like ClassNameUDP And NameUdp Like NameFieldUDP Then
                    If TypeOf IdUdp Is UDPDouble Then 'вещественное с двойной точностью
                        Try
                            Dim UDp As UDPDouble = TryCast(IdUdp, UDPDouble)
                            Dim valDbl As Double = CDbl(UDPval)
                            PointCOGO.SetUDPValue(UDp, valDbl)
                            FuncWriteUDPValue = True
                            Exit For
                        Catch ex As System.InvalidOperationException
                        End Try
                    ElseIf TypeOf IdUdp Is UDPBoolean Then 'false/true
                        Try
                            Dim UDp As UDPBoolean = TryCast(IdUdp, UDPBoolean)
                            Dim valDbl As Boolean = CBool(UDPval)
                            PointCOGO.SetUDPValue(UDp, valDbl)
                            FuncWriteUDPValue = True
                            Exit For
                        Catch ex As System.InvalidOperationException
                        End Try
                    ElseIf TypeOf IdUdp Is UDPEnumeration Then '
                        Try
                            Dim UDp As UDPEnumeration = TryCast(IdUdp, UDPEnumeration)
                        Catch ex As System.InvalidOperationException
                        End Try
                    ElseIf TypeOf IdUdp Is UDPInteger Then 'целое
                        Try
                            Dim UDp As UDPInteger = TryCast(IdUdp, UDPInteger)
                            Dim valDbl As Integer = CInt(UDPval)
                            PointCOGO.SetUDPValue(UDp, valDbl)
                            FuncWriteUDPValue = True
                            Exit For
                        Catch ex As System.InvalidOperationException
                        End Try
                    ElseIf TypeOf IdUdp Is UDPString Then 'строка
                        Try
                            Dim UDp As UDPString = TryCast(IdUdp, UDPString)
                            PointCOGO.SetUDPValue(UDp, UDPval)
                            FuncWriteUDPValue = True
                            Exit For
                        Catch ex As System.InvalidOperationException
                        End Try
                    End If
                End If
            End If
        Next
    End Function



End Class
