Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Public Class FuncAlignment
    '==================================================================================================================================================================================================
    '==================================================================================================================================================================================================
    'характеристики элементов трассы
    '(0-ПК НАЧАЛА, 1-ПК КОНЦА, 2 - X центра дуги (SPI для спирали), 3 - Y центра дуги (SPI для спирали), 4- тип объекта (0-линия,1-дуга,2-спираль), 5 - X начала элемента, 6 - Y начала элемента, 7 - X конца элемента, 8 - Y конца элемента, 9 - пк SPI или центра дуги, 10 - радиус дуги или спирали (входящий),11 - радиус спирали (выходящий), 12 - длина объекта, 13 - направление объекта (дир.угол), 14 - точка О спирали Х, 15- точка О спирали Y
    '==================================================================================================================================================================================================
    '==================================================================================================================================================================================================
    'объект линия
    Public Shared Function FuncCharacteristicsLine(ByVal align As Alignment, ByVal AlignLine As AlignmentLine, ByRef ArrayEntity As Double(,), ByVal PKStartAlign As Double, ByVal PKEndAlign As Double) As Boolean
        FuncCharacteristicsLine = False
        'если сегмент нулевой длины то не обрабатываем его и выходим из функции
        If AlignLine.Length = 0 Then Exit Function
        Dim PKStartObj As Double = AlignLine.StartStation 'пикет начала линии 
        Dim PKEndObj As Double = AlignLine.EndStation 'пикет конца линии
        Dim count As Integer = 0 'число элементов в массиве
        'определяем число элементов в массиве
        If IsArray(ArrayEntity) = True Then
            count = ArrayEntity.GetUpperBound(1) + 1
        End If
        'условие до точки старта
        If PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj <= PKStartAlign And PKEndObj < PKEndAlign Then
            Exit Function
        End If
        If PKStartObj > PKStartAlign And PKStartObj >= PKEndAlign And PKEndObj > PKStartAlign And PKEndObj > PKEndAlign Then
            Exit Function
        End If
        If PKStartObj >= PKStartAlign Then
            If PKStartObj < PKEndAlign Then
                If PKEndObj > PKStartObj Then
                    If PKEndObj <= PKEndAlign Then
                        Dim a = 1
                    End If
                End If
            End If
        End If
        If PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignLine.StartStation 'начальный пикет
            ArrayEntity(1, count) = AlignLine.EndStation 'конечный пикет
            ArrayEntity(4, count) = 0 'тип объекта
            ArrayEntity(5, count) = AlignLine.StartPoint.X 'координаты
            ArrayEntity(6, count) = AlignLine.StartPoint.Y
            ArrayEntity(7, count) = AlignLine.EndPoint.X
            ArrayEntity(8, count) = AlignLine.EndPoint.Y
            ArrayEntity(12, count) = AlignLine.Length 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта (дирекционный угол)
            count += 1
            'объект выходящим концом находится внутри диапазона
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign
            ArrayEntity(1, count) = AlignLine.EndStation
            ArrayEntity(4, count) = 0
            Dim X As Double = -1
            Dim Y As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                ArrayEntity(5, count) = X
                ArrayEntity(6, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(0, count) = AlignLine.StartStation
                ArrayEntity(5, count) = AlignLine.StartPoint.X
                ArrayEntity(6, count) = AlignLine.StartPoint.Y
            End Try
            ArrayEntity(7, count) = AlignLine.EndPoint.X
            ArrayEntity(8, count) = AlignLine.EndPoint.Y
            ArrayEntity(12, count) = AlignLine.EndStation - PKStartAlign 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта
            count += 1
            'объект входящим концом находится внутри диапазона
        ElseIf (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignLine.StartStation
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 0
            ArrayEntity(5, count) = AlignLine.StartPoint.X
            ArrayEntity(6, count) = AlignLine.StartPoint.Y
            Dim X As Double = -1
            Dim Y As Double = -1
            Try
                align.PointLocation(PKEndAlign, 0, X, Y)
                ArrayEntity(7, count) = X
                ArrayEntity(8, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(1, count) = AlignLine.EndStation
                ArrayEntity(7, count) = AlignLine.EndPoint.X
                ArrayEntity(8, count) = AlignLine.EndPoint.Y
            End Try
            ArrayEntity(12, count) = PKEndAlign - AlignLine.StartStation 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта
            'объект обоими концами вне диапазона но середина объекта внутри
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 0
            Dim X As Double = -1
            Dim Y As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                ArrayEntity(5, count) = X
                ArrayEntity(6, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(0, count) = AlignLine.StartStation
                ArrayEntity(5, count) = AlignLine.StartPoint.X
                ArrayEntity(6, count) = AlignLine.StartPoint.Y
            End Try
            X = -1
            Y = -1
            Try
                align.PointLocation(PKEndAlign, 0, X, Y)
                ArrayEntity(7, count) = X
                ArrayEntity(8, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(1, count) = AlignLine.EndStation
                ArrayEntity(7, count) = AlignLine.EndPoint.X
                ArrayEntity(8, count) = AlignLine.EndPoint.Y
            End Try
            ArrayEntity(12, count) = PKEndAlign - PKEndAlign 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта
        Else
            FuncCharacteristicsLine = False
            Exit Function
        End If
        FuncCharacteristicsLine = True
    End Function
    'объект вложенная линия
    Public Shared Function FuncCharacteristicsSubLine(ByVal align As Alignment, ByVal AlignLine As AlignmentSubEntityLine, ByRef ArrayEntity As Double(,), ByVal PKStartAlign As Double, ByVal PKEndAlign As Double) As Boolean
        FuncCharacteristicsSubLine = False
        If AlignLine.Length = 0 Then Exit Function
        Dim PKStartObj As Double = AlignLine.StartStation 'пикет начала линии 
        Dim PKEndObj As Double = AlignLine.EndStation 'пикет конца линии
        Dim count As Integer = 0 'число элементов в массиве
        'определяем число элементов в массиве
        'условие до точки старта
        If PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj <= PKStartAlign And PKEndObj < PKEndAlign Then
            Exit Function
        End If
        If PKStartObj > PKStartAlign And PKStartObj >= PKEndAlign And PKEndObj > PKStartAlign And PKEndObj > PKEndAlign Then
            Exit Function
        End If
        If IsArray(ArrayEntity) = True Then
            count = ArrayEntity.GetUpperBound(1) + 1
        End If
        If (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignLine.StartStation 'начальный пикет
            ArrayEntity(1, count) = AlignLine.EndStation 'конечный пикет
            ArrayEntity(4, count) = 0 'тип объекта
            ArrayEntity(5, count) = AlignLine.StartPoint.X 'координаты
            ArrayEntity(6, count) = AlignLine.StartPoint.Y
            ArrayEntity(7, count) = AlignLine.EndPoint.X
            ArrayEntity(8, count) = AlignLine.EndPoint.Y
            ArrayEntity(12, count) = AlignLine.Length 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта (дирекционный угол)
            count += 1
            'объект выходящим концом находится внутри диапазона
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign
            ArrayEntity(1, count) = AlignLine.EndStation
            ArrayEntity(4, count) = 0
            Dim X As Double = -1
            Dim Y As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                ArrayEntity(5, count) = X
                ArrayEntity(6, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(0, count) = AlignLine.StartStation
                ArrayEntity(5, count) = AlignLine.StartPoint.X
                ArrayEntity(6, count) = AlignLine.StartPoint.Y
            End Try
            ArrayEntity(7, count) = AlignLine.EndPoint.X
            ArrayEntity(8, count) = AlignLine.EndPoint.Y
            ArrayEntity(12, count) = AlignLine.EndStation - PKStartAlign 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта
            count += 1
            'объект входящим концом находится внутри диапазона
        ElseIf (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignLine.StartStation
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 0
            ArrayEntity(5, count) = AlignLine.StartPoint.X
            ArrayEntity(6, count) = AlignLine.StartPoint.Y
            Dim X As Double = -1
            Dim Y As Double = -1
            Try
                align.PointLocation(PKEndAlign, 0, X, Y)
                ArrayEntity(7, count) = X
                ArrayEntity(8, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(1, count) = AlignLine.EndStation
                ArrayEntity(7, count) = AlignLine.EndPoint.X
                ArrayEntity(8, count) = AlignLine.EndPoint.Y
            End Try
            ArrayEntity(12, count) = PKEndAlign - AlignLine.StartStation 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта
            'объект обоими концами вне диапазона но середина объекта внутри
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 0
            Dim X As Double = -1
            Dim Y As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                ArrayEntity(5, count) = X
                ArrayEntity(6, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(0, count) = AlignLine.StartStation
                ArrayEntity(5, count) = AlignLine.StartPoint.X
                ArrayEntity(6, count) = AlignLine.StartPoint.Y
            End Try
            X = -1
            Y = -1
            Try
                align.PointLocation(PKEndAlign, 0, X, Y)
                ArrayEntity(7, count) = X
                ArrayEntity(8, count) = Y
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                ArrayEntity(1, count) = AlignLine.EndStation
                ArrayEntity(7, count) = AlignLine.EndPoint.X
                ArrayEntity(8, count) = AlignLine.EndPoint.Y
            End Try
            ArrayEntity(12, count) = PKEndAlign - PKEndAlign 'длина объекта
            ArrayEntity(13, count) = AlignLine.Direction 'направление объекта
        End If
        FuncCharacteristicsSubLine = True
    End Function
    'объект дуга
    Public Shared Function FuncCharacteristicsArc(ByVal align As Alignment, ByVal AlignArc As AlignmentArc, ByRef ArrayEntity As Double(,), ByVal PKStartAlign As Double, ByVal PKEndAlign As Double) As Boolean
        FuncCharacteristicsArc = False
        If AlignArc.Length = 0 Then
            Exit Function
        End If
        Dim PKStartObj As Double = AlignArc.StartStation 'пикет начала дуги
        Dim PKEndObj As Double = AlignArc.EndStation 'пикет конца дуги
        Dim count As Integer = 0 'число элементов в массиве
        'определяем число элементов в массиве
        'условие до точки старта
        If PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj <= PKStartAlign And PKEndObj < PKEndAlign Then
            Exit Function
        End If
        If PKStartObj > PKStartAlign And PKStartObj >= PKEndAlign And PKEndObj > PKStartAlign And PKEndObj > PKEndAlign Then
            Exit Function
        End If
        If IsArray(ArrayEntity) = True Then
            count = ArrayEntity.GetUpperBound(1) + 1
        End If
        If (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            'смотрим длину дуги
            Dim LArc As Double = AlignArc.Length
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation 'пикет начала
                ArrayEntity(1, count) = AlignArc.EndStation 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = AlignArc.StartPoint.X 'координаты начала дуги
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                ArrayEntity(7, count) = AlignArc.EndPoint.X 'координаты конца дуги
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.Length
                count += 1
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation 'пикет начала
                Dim PK As Double = (AlignArc.StartStation + AlignArc.EndStation) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = AlignArc.StartPoint.X 'координаты начала дуги
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    Dim MiddleArc As Point2d = MathFunction.FuncFindMiddleArc(AlignArc.StartPoint, AlignArc.EndPoint, AlignArc.CenterPoint)
                    Xsr = MiddleArc.X
                    Ysr = MiddleArc.Y
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.Length / 2
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = AlignArc.EndStation 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = AlignArc.EndPoint.X 'координаты конца дуги
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.Length / 2
                count += 1
            End If
            'объект выходящим концом находится внутри диапазона
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            '================================================================================
            'вычисляем координаты нового начала дуги
            Dim Xn As Double = -1
            Dim Yn As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, Xn, Yn)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '===============================================================================
            'смотрим длину дуги
            Dim LArc As Double = AlignArc.EndStation - PKStartAlign
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign
                ArrayEntity(1, count) = AlignArc.EndStation
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1
                ArrayEntity(5, count) = Xn
                ArrayEntity(6, count) = Yn
                ArrayEntity(7, count) = AlignArc.EndPoint.X
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.EndStation - PKStartAlign 'длина объекта
                count += 1
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign 'пикет начала
                Dim PK As Double = (PKStartAlign + AlignArc.EndStation) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xn 'координаты начала дуги
                ArrayEntity(6, count) = Yn
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PK - PKStartAlign
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = AlignArc.EndStation 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = AlignArc.EndPoint.X 'координаты конца дуги
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.EndStation - PK
                count += 1
            End If
            'объект входящим концом находится внутри диапазона
        ElseIf (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            '================================================================================
            'вычисляем координаты нового конца дуги
            Dim Xk As Double = -1
            Dim Yk As Double = -1
            Try
                align.PointLocation(PKEndAlign, 0, Xk, Yk)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '===============================================================================
            'смотрим длину дуги
            Dim LArc As Double = PKEndAlign - AlignArc.StartStation
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation
                ArrayEntity(1, count) = PKEndAlign
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1
                ArrayEntity(5, count) = AlignArc.StartPoint.X
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                ArrayEntity(7, count) = Xk
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - AlignArc.StartStation 'длина объекта
                count += 1
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation 'пикет начала
                Dim PK As Double = (AlignArc.StartStation + PKEndAlign) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = AlignArc.StartPoint.X 'координаты начала дуги
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PK - AlignArc.StartStation
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = PKEndAlign 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = Xk 'координаты конца дуги
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - PK
                count += 1
            End If
            'объект обоими концами вне диапазона но середина объекта внутри
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            '================================================================================
            'координаты точки начала
            Dim Xn As Double = -1
            Dim Yn As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, Xn, Yn)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '================================================================================
            'координаты точки конца
            Dim Xk As Double = -1
            Dim Yk As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, Xk, Yk)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '===============================================================================
            'смотрим длину дуги
            Dim LArc As Double = PKEndAlign - PKStartAlign
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign
                ArrayEntity(1, count) = PKEndAlign
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1
                ArrayEntity(5, count) = Xn
                ArrayEntity(6, count) = Yn
                ArrayEntity(7, count) = Xk
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - PKEndAlign 'длина объекта
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign 'пикет начала
                Dim PK As Double = (PKStartAlign + PKEndAlign) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xn 'координаты начала дуги
                ArrayEntity(6, count) = Yn
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PK - PKStartAlign
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = PKEndAlign 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = Xk 'координаты конца дуги
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - PK
                count += 1
            End If
        End If
        FuncCharacteristicsArc = True
    End Function
    'объект вложенная дуга
    Public Shared Function FuncCharacteristicsSubArc(ByVal align As Alignment, ByVal AlignArc As AlignmentSubEntityArc, ByRef ArrayEntity As Double(,), ByVal PKStartAlign As Double, ByVal PKEndAlign As Double) As Boolean
        FuncCharacteristicsSubArc = False
        If AlignArc.Length = 0 Then
            Exit Function
        End If
        Dim PKStartObj As Double = AlignArc.StartStation 'пикет начала дуги
        Dim PKEndObj As Double = AlignArc.EndStation 'пикет конца дуги
        Dim count As Integer = 0 'число элементов в массиве
        'определяем число элементов в массиве
        'условие до точки старта
        If PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj <= PKStartAlign And PKEndObj < PKEndAlign Then
            Exit Function
        End If
        If PKStartObj > PKStartAlign And PKStartObj >= PKEndAlign And PKEndObj > PKStartAlign And PKEndObj > PKEndAlign Then
            Exit Function
        End If
        If IsArray(ArrayEntity) = True Then
            count = ArrayEntity.GetUpperBound(1) + 1
        End If
        If (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            'смотрим длину дуги
            Dim LArc As Double = AlignArc.Length
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation 'пикет начала
                ArrayEntity(1, count) = AlignArc.EndStation 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = AlignArc.StartPoint.X 'координаты начала дуги
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                ArrayEntity(7, count) = AlignArc.EndPoint.X 'координаты конца дуги
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.Length
                count += 1
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation 'пикет начала
                Dim PK As Double = (AlignArc.StartStation + AlignArc.EndStation) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = AlignArc.StartPoint.X 'координаты начала дуги
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    Dim MiddleArc As Point2d = MathFunction.FuncFindMiddleArc(AlignArc.StartPoint, AlignArc.EndPoint, AlignArc.CenterPoint)
                    Xsr = MiddleArc.X
                    Ysr = MiddleArc.Y
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.Length / 2
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = AlignArc.EndStation 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = AlignArc.EndPoint.X 'координаты конца дуги
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.Length / 2
                count += 1
            End If
            'объект выходящим концом находится внутри диапазона
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            '================================================================================
            'вычисляем координаты нового начала дуги
            Dim Xn As Double = -1
            Dim Yn As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, Xn, Yn)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '===============================================================================
            'смотрим длину дуги
            Dim LArc As Double = AlignArc.EndStation - PKStartAlign
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign
                ArrayEntity(1, count) = AlignArc.EndStation
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1
                ArrayEntity(5, count) = Xn
                ArrayEntity(6, count) = Yn
                ArrayEntity(7, count) = AlignArc.EndPoint.X
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.EndStation - PKStartAlign 'длина объекта
                count += 1
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign 'пикет начала
                Dim PK As Double = (PKStartAlign + AlignArc.EndStation) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xn 'координаты начала дуги
                ArrayEntity(6, count) = Yn
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PK - PKStartAlign
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = AlignArc.EndStation 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = AlignArc.EndPoint.X 'координаты конца дуги
                ArrayEntity(8, count) = AlignArc.EndPoint.Y
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = AlignArc.EndStation - PK
                count += 1
            End If
            'объект входящим концом находится внутри диапазона
        ElseIf (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            '================================================================================
            'вычисляем координаты нового конца дуги
            Dim Xk As Double = -1
            Dim Yk As Double = -1
            Try
                align.PointLocation(PKEndAlign, 0, Xk, Yk)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '===============================================================================
            'смотрим длину дуги
            Dim LArc As Double = PKEndAlign - AlignArc.StartStation
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation
                ArrayEntity(1, count) = PKEndAlign
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1
                ArrayEntity(5, count) = AlignArc.StartPoint.X
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                ArrayEntity(7, count) = Xk
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - AlignArc.StartStation 'длина объекта
                count += 1
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = AlignArc.StartStation 'пикет начала
                Dim PK As Double = (AlignArc.StartStation + PKEndAlign) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = AlignArc.StartPoint.X 'координаты начала дуги
                ArrayEntity(6, count) = AlignArc.StartPoint.Y
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PK - AlignArc.StartStation
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = PKEndAlign 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = Xk 'координаты конца дуги
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - PK
                count += 1
            End If
            'объект обоими концами вне диапазона но середина объекта внутри
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            '================================================================================
            'координаты точки начала
            Dim Xn As Double = -1
            Dim Yn As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, Xn, Yn)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '================================================================================
            'координаты точки конца
            Dim Xk As Double = -1
            Dim Yk As Double = -1
            Try
                align.PointLocation(PKStartAlign, 0, Xk, Yk)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            '===============================================================================
            'смотрим длину дуги
            Dim LArc As Double = PKEndAlign - PKStartAlign
            'смотрим радиус дуги
            Dim RArc As Double = AlignArc.Radius
            'считаем длину огружности (дополняем дугу до окружности)
            Dim LCirc As Double = 2 * Math.PI * RArc
            'сравниваем длины
            Dim DeldaL As Double = Math.Round(LCirc / 2 - LArc, 3)
            If DeldaL >= 0 Then
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign
                ArrayEntity(1, count) = PKEndAlign
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1
                ArrayEntity(5, count) = Xn
                ArrayEntity(6, count) = Yn
                ArrayEntity(7, count) = Xk
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - PKEndAlign 'длина объекта
            Else
                'делим дугу пополам
                'находим новое пикетажное положение конца дуги (оно же начало 2 дуги)
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PKStartAlign 'пикет начала
                Dim PK As Double = (PKStartAlign + PKEndAlign) / 2
                ArrayEntity(1, count) = PK 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xn 'координаты начала дуги
                ArrayEntity(6, count) = Yn
                'находим координаты конца криыой
                Dim Xsr As Double = -1
                Dim Ysr As Double = -1
                Try
                    align.PointLocation(PK, 0, Xsr, Ysr)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                End Try
                ArrayEntity(7, count) = Xsr 'координаты конца дуги
                ArrayEntity(8, count) = Ysr
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PK - PKStartAlign
                count += 1
                '================================================================================================
                ReDim Preserve ArrayEntity(15, count)
                ArrayEntity(0, count) = PK 'пикет начала
                ArrayEntity(1, count) = PKEndAlign 'пикет конца
                ArrayEntity(2, count) = AlignArc.CenterPoint.X 'координаты центра дуги
                ArrayEntity(3, count) = AlignArc.CenterPoint.Y
                ArrayEntity(4, count) = 1 'тип объекта
                ArrayEntity(5, count) = Xsr 'координаты начала дуги
                ArrayEntity(6, count) = Ysr
                ArrayEntity(7, count) = Xk 'координаты конца дуги
                ArrayEntity(8, count) = Yk
                Try
                    ArrayEntity(9, count) = AlignArc.PIStation
                Catch ex As InvalidOperationException
                    ArrayEntity(9, count) = -100000 'ошибочное значение
                End Try
                ArrayEntity(10, count) = AlignArc.Radius
                ArrayEntity(11, count) = AlignArc.Radius
                ArrayEntity(12, count) = PKEndAlign - PK
                count += 1
            End If
        End If
        FuncCharacteristicsSubArc = True
    End Function
    'объект спираль
    Public Shared Function FuncCharacteristicsSpiral(ByVal align As Alignment, ByVal AlignSpiral As AlignmentSpiral, ByRef ArrayEntity As Double(,), ByVal PKStartAlign As Double, ByVal PKEndAlign As Double) As Boolean
        FuncCharacteristicsSpiral = False
        If AlignSpiral.Length = 0 Then
            Exit Function
        End If
        Dim PKStartObj As Double = AlignSpiral.StartStation 'пикет начала дуги
        Dim PKEndObj As Double = AlignSpiral.EndStation 'пикет конца дуги
        Dim count As Integer = 0 'число элементов в массиве
        'определяем число элементов в массиве
        'условие до точки старта
        If PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj <= PKStartAlign And PKEndObj < PKEndAlign Then
            Exit Function
        End If
        If PKStartObj > PKStartAlign And PKStartObj >= PKEndAlign And PKEndObj > PKStartAlign And PKEndObj > PKEndAlign Then
            Exit Function
        End If
        If IsArray(ArrayEntity) = True Then
            count = ArrayEntity.GetUpperBound(1) + 1
        End If
        If (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignSpiral.StartStation 'пикет начала
            ArrayEntity(1, count) = AlignSpiral.EndStation 'пикет конца
            ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'координаты SPI точки
            ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
            ArrayEntity(4, count) = 2 'тип объекта
            ArrayEntity(5, count) = AlignSpiral.StartPoint.X 'координаты начала спирали
            ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
            ArrayEntity(7, count) = AlignSpiral.EndPoint.X 'координаты конца спирали
            ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
            ArrayEntity(9, count) = AlignSpiral.SPIStation 'пикет SPI точки
            Dim RIn As Double = AlignSpiral.RadiusIn
            If Double.IsInfinity(RIn) Then
                ArrayEntity(10, count) = -1 'входящий радиус
            Else
                ArrayEntity(10, count) = AlignSpiral.RadiusIn 'входящий радиус
            End If
            Dim ROut As Double = AlignSpiral.RadiusOut
            If Double.IsInfinity(ROut) Then
                ArrayEntity(11, count) = -1 'входящий радиус
            Else
                ArrayEntity(11, count) = AlignSpiral.RadiusOut 'входящий радиус
            End If
            ArrayEntity(12, count) = AlignSpiral.Length 'длина спирали
            ArrayEntity(14, count) = AlignSpiral.RadialPoint.X 'точка О спирали
            ArrayEntity(15, count) = AlignSpiral.RadialPoint.Y 'точка О спирали
            count += 1
            'объект выходящим концом находится внутри диапазона
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign 'начало переходной кривой (ПК)
            ArrayEntity(1, count) = AlignSpiral.EndStation 'конец переходной кривой (ПК)
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim X As Double = -1 'начальные значения точки начала
            Dim Y As Double = -1
            Dim XY As Point2d = Nothing
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                XY = New Point2d(X, Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1 As Double = -1
            Dim Y1 As Double = -1
            Dim XY1 As Point2d = Nothing
            Dim PKPrev As Double = PKStartAlign - 0.001
            Try
                align.PointLocation(PKPrev, 0, X1, Y1)
                XY1 = New Point2d(X1, Y1)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2 As Double = -1
            Dim Y2 As Double = -1
            Dim XY2 As Point2d = Nothing
            Dim PKNext As Double = PKStartAlign + 0.001
            Try
                align.PointLocation(PKNext, 0, X2, Y2)
                XY2 = New Point2d(X2, Y2)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1 As Vector2d = Nothing
            Dim Vect2 As Vector2d = Nothing
            Dim Dir1 As Double = -1
            Dim Dir2 As Double = -1
            If X1 <> -1 And Y1 <> -1 And X <> -1 And Y <> -1 Then
                Vect1 = XY1.GetVectorTo(XY)
                Dir1 = Vect1.Angle
            End If
            If X2 <> -1 And Y2 <> -1 And X <> -1 And Y <> -1 Then
                Vect2 = XY.GetVectorTo(XY2)
                Dir2 = Vect2.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dir As Double = -1
            If Dir1 > -1 And Dir2 > -1 Then
                Dir = (Dir1 + Dir2) / 2
            ElseIf Dir > -1 And Dir2 = -1 Then
                Dir = Dir1
            ElseIf Dir1 = -1 And Dir2 > -1 Then
                Dir = Dir2
            End If
            'определяем новые координаты для направления первого вектора
            If Dir > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.SPIStation - AlignSpiral.StartStation)
                Dim Xv As Double = X + LenghtSPI * Math.Cos(Dir)
                Dim Yv As Double = Y + LenghtSPI * Math.Sin(Dir)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(X, Y, 0)
                Dim LineV1 As Line = New Line(XY13D, XYv3D)
                Dim XY23D As Point3d = New Point3d(AlignSpiral.EndPoint.X, AlignSpiral.EndPoint.Y, 0)
                Dim XYSPI3D As Point3d = New Point3d(AlignSpiral.SPIPoint.X, AlignSpiral.SPIPoint.Y, 0)
                Dim LineV2 As Line = New Line(XY23D, XYSPI3D)
                'ищем точку пересечения линий
                Dim ColPoint As Point3dCollection = New Point3dCollection
                LineV1.IntersectWith(LineV2, Intersect.ExtendBoth, ColPoint, IntPtr.Zero, IntPtr.Zero)
                If ColPoint.Count > 0 Then
                    For Each pt3d As Point3d In ColPoint
                        ArrayEntity(2, count) = pt3d.X 'новые координаты SPI точки для спирали
                        ArrayEntity(3, count) = pt3d.Y
                        ArrayEntity(5, count) = X
                        ArrayEntity(6, count) = Y
                        Exit For
                    Next
                End If
            Else
                ArrayEntity(0, count) = AlignSpiral.StartStation 'пикет начала сегмента
                ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'новые координаты SPI точки для спирали
                ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
                ArrayEntity(5, count) = AlignSpiral.StartPoint.X
                ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
            End If
            ArrayEntity(4, count) = 2
            ArrayEntity(7, count) = AlignSpiral.EndPoint.X
            ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
            ArrayEntity(9, count) = AlignSpiral.SPIStation
            'перезаписываем входящий радиус
            Dim S As Double = AlignSpiral.EndStation - PKStartAlign 'длина сегмента от начала конструкции до точки выбора
            Dim R As Double = AlignSpiral.RadiusOut
            'если все как положено входящий радиус бесконечен а выходящий имеет число
            If Double.IsInfinity(R) = False Then
                ArrayEntity(10, count) = (R * AlignSpiral.Length) / S
                ArrayEntity(11, count) = R
            Else
                'маловероятно но делаем все наоборот
                R = AlignSpiral.RadiusIn
                ArrayEntity(10, count) = R
                ArrayEntity(11, count) = (R * AlignSpiral.Length) / S
            End If
            ArrayEntity(12, count) = S 'длина объекта
            count += 1
            'объект входящим концом находится внутри диапазона
        ElseIf (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignSpiral.StartStation
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 2
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim X As Double = -1 'начальные значения точки конца
            Dim Y As Double = -1
            Dim XY As Point2d = Nothing
            Try
                align.PointLocation(PKEndAlign, 0, X, Y)
                XY = New Point2d(X, Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1 As Double = -1
            Dim Y1 As Double = -1
            Dim XY1 As Point2d = Nothing
            Dim PKPrev As Double = PKEndAlign - 0.001
            Try
                align.PointLocation(PKPrev, 0, X1, Y1)
                XY1 = New Point2d(X1, Y1)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2 As Double = -1
            Dim Y2 As Double = -1
            Dim XY2 As Point2d = Nothing
            Dim PKNext As Double = PKEndAlign + 0.001
            Try
                align.PointLocation(PKNext, 0, X2, Y2)
                XY2 = New Point2d(X2, Y2)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1 As Vector2d = Nothing
            Dim Vect2 As Vector2d = Nothing
            Dim Dir1 As Double = -1
            Dim Dir2 As Double = -1
            If X1 <> -1 And Y1 <> -1 And X <> -1 And Y <> -1 Then
                Vect1 = XY.GetVectorTo(XY1)
                Dir1 = Vect1.Angle
            End If
            If X2 <> -1 And Y2 <> -1 And X <> -1 And Y <> -1 Then
                Vect2 = XY2.GetVectorTo(XY)
                Dir2 = Vect2.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dir As Double = -1
            If Dir1 > -1 And Dir2 > -1 Then
                Dir = (Dir1 + Dir2) / 2
            ElseIf Dir > -1 And Dir2 = -1 Then
                Dir = Dir1
            ElseIf Dir1 = -1 And Dir2 > -1 Then
                Dir = Dir2
            End If
            'определяем новые координаты для направления первого вектора
            If Dir > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.EndStation - AlignSpiral.SPIStation)
                Dim Xv As Double = X + LenghtSPI * Math.Cos(Dir)
                Dim Yv As Double = Y + LenghtSPI * Math.Sin(Dir)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(X, Y, 0)
                Dim LineV1 As Line = New Line(XY13D, XYv3D)
                Dim XY23D As Point3d = New Point3d(AlignSpiral.StartPoint.X, AlignSpiral.StartPoint.Y, 0)
                Dim XYSPI3D As Point3d = New Point3d(AlignSpiral.SPIPoint.X, AlignSpiral.SPIPoint.Y, 0)
                Dim LineV2 As Line = New Line(XY23D, XYSPI3D)
                'ищем точку пересечения линий
                Dim ColPoint As Point3dCollection = New Point3dCollection
                LineV1.IntersectWith(LineV2, Intersect.ExtendBoth, ColPoint, IntPtr.Zero, IntPtr.Zero)
                If ColPoint.Count > 0 Then
                    For Each pt3d As Point3d In ColPoint
                        ArrayEntity(2, count) = pt3d.X 'новые координаты SPI точки для спирали
                        ArrayEntity(3, count) = pt3d.Y
                        ArrayEntity(7, count) = X
                        ArrayEntity(8, count) = Y
                        Exit For
                    Next
                End If
            Else
                ArrayEntity(1, count) = AlignSpiral.EndStation 'пикет начала сегмента
                ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'новые координаты SPI точки для спирали
                ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
                ArrayEntity(7, count) = AlignSpiral.EndPoint.X
                ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
            End If

            ArrayEntity(5, count) = AlignSpiral.StartPoint.X
            ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
            ArrayEntity(9, count) = AlignSpiral.SPIStation
            'перезаписываем входящий радиус
            Dim S As Double = PKEndAlign - AlignSpiral.StartStation 'длина сегмента от начала конструкции до точки выбора
            Dim R As Double = AlignSpiral.RadiusIn
            'если все как положено входящий радиус бесконечен а выходящий имеет число
            If Double.IsInfinity(R) = False Then
                ArrayEntity(10, count) = R
                ArrayEntity(11, count) = (R * AlignSpiral.Length) / S
            Else
                'маловероятно но делаем все наоборот
                R = AlignSpiral.RadiusOut
                ArrayEntity(10, count) = (R * AlignSpiral.Length) / S
                ArrayEntity(11, count) = R
            End If
            ArrayEntity(12, count) = S 'длина объекта
            count += 1
            'объект обоими концами вне диапазона но середина объекта внутри
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 2
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim X As Double = -1 'начальные значения точки начала
            Dim Y As Double = -1
            Dim XY As Point2d = Nothing
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                XY = New Point2d(X, Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1 As Double = -1
            Dim Y1 As Double = -1
            Dim XY1 As Point2d = Nothing
            Dim PKPrev As Double = PKStartAlign - 0.001
            Try
                align.PointLocation(PKPrev, 0, X1, Y1)
                XY1 = New Point2d(X1, Y1)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2 As Double = -1
            Dim Y2 As Double = -1
            Dim XY2 As Point2d = Nothing
            Dim PKNext As Double = PKStartAlign + 0.001
            Try
                align.PointLocation(PKNext, 0, X2, Y2)
                XY2 = New Point2d(X2, Y2)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1 As Vector2d = Nothing
            Dim Vect2 As Vector2d = Nothing
            Dim Dir1 As Double = -1
            Dim Dir2 As Double = -1
            If X1 <> -1 And Y1 <> -1 And X <> -1 And Y <> -1 Then
                Vect1 = XY1.GetVectorTo(XY)
                Dir1 = Vect1.Angle
            End If
            If X2 <> -1 And Y2 <> -1 And X <> -1 And Y <> -1 Then
                Vect2 = XY.GetVectorTo(XY2)
                Dir2 = Vect2.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dir As Double = -1
            If Dir1 > -1 And Dir2 > -1 Then
                Dir = (Dir1 + Dir2) / 2
            ElseIf Dir > -1 And Dir2 = -1 Then
                Dir = Dir1
            ElseIf Dir1 = -1 And Dir2 > -1 Then
                Dir = Dir2
            End If
            'определяем новые координаты для направления первого вектора
            Dim LineV1 As Line = New Line()
            If Dir > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.SPIStation - AlignSpiral.StartStation)
                Dim Xv As Double = X + LenghtSPI * Math.Cos(Dir)
                Dim Yv As Double = Y + LenghtSPI * Math.Sin(Dir)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(X, Y, 0)
                LineV1 = New Line(XY13D, XYv3D)
            End If
            '=========================================================================================
            'второй конец
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim Xk As Double = -1 'начальные значения точки конца
            Dim Yk As Double = -1
            Dim XYk As Point2d = Nothing
            Try
                align.PointLocation(PKEndAlign, 0, Xk, Yk)
                XYk = New Point2d(Xk, Yk)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1k As Double = -1
            Dim Y1k As Double = -1
            Dim XY1k As Point2d = Nothing
            Dim PKPrevk As Double = PKEndAlign - 0.001
            Try
                align.PointLocation(PKPrevk, 0, X1k, Y1k)
                XY1k = New Point2d(X1k, Y1k)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2k As Double = -1
            Dim Y2k As Double = -1
            Dim XY2k As Point2d = Nothing
            Dim PKNextk As Double = PKEndObj + 0.001
            Try
                align.PointLocation(PKNextk, 0, X2k, Y2k)
                XY2k = New Point2d(X2k, Y2k)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1k As Vector2d = Nothing
            Dim Vect2k As Vector2d = Nothing
            Dim Dir1k As Double = -1
            Dim Dir2k As Double = -1
            If X1k <> -1 And Y1k <> -1 And Xk <> -1 And Yk <> -1 Then
                Vect1k = XYk.GetVectorTo(XY1k)
                Dir1k = Vect1k.Angle
            End If
            If X2k <> -1 And Y2k <> -1 And Xk <> -1 And Yk <> -1 Then
                Vect2k = XY2k.GetVectorTo(XYk)
                Dir2k = Vect2k.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dirk As Double = -1
            If Dir1k > -1 And Dir2k > -1 Then
                Dirk = (Dir1k + Dir2k) / 2
            ElseIf Dirk > -1 And Dir2k = -1 Then
                Dirk = Dir1k
            ElseIf Dir1k = -1 And Dir2k > -1 Then
                Dirk = Dir2k
            End If
            'определяем новые координаты для направления первого вектора
            Dim LineV2 As Line = New Line()
            If Dirk > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.EndStation - AlignSpiral.SPIStation)
                Dim Xv As Double = Xk + LenghtSPI * Math.Cos(Dirk)
                Dim Yv As Double = Yk + LenghtSPI * Math.Sin(Dirk)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(Xk, Yk, 0)
                LineV2 = New Line(XY13D, XYv3D)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем точки пересечения
            Dim ColPoint As Point3dCollection = New Point3dCollection
            LineV1.IntersectWith(LineV2, Intersect.ExtendBoth, ColPoint, IntPtr.Zero, IntPtr.Zero)
            If ColPoint.Count > 0 Then
                For Each pt3d As Point3d In ColPoint
                    ArrayEntity(2, count) = pt3d.X 'новые координаты SPI точки для спирали
                    ArrayEntity(3, count) = pt3d.Y
                    ArrayEntity(5, count) = X
                    ArrayEntity(6, count) = Y
                    ArrayEntity(7, count) = Xk
                    ArrayEntity(8, count) = Yk
                    ArrayEntity(12, count) = PKEndAlign - PKStartAlign 'длина объекта
                    Exit For
                Next
            Else
                ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'новые координаты SPI точки для спирали
                ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
                ArrayEntity(5, count) = AlignSpiral.StartPoint.X
                ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
                ArrayEntity(7, count) = AlignSpiral.EndPoint.X
                ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
                ArrayEntity(12, count) = AlignSpiral.EndStation - AlignSpiral.StartStation 'длина объекта
            End If
            'вычисляем радиусы
            Dim R As Double = AlignSpiral.RadiusIn
            Dim L As Double = AlignSpiral.Length
            If Double.IsInfinity(R) = False Then
                Dim S1 As Double = PKStartAlign - AlignSpiral.StartStation
                Dim S2 As Double = PKEndAlign - AlignSpiral.EndStation
                ArrayEntity(10, count) = (R * L) / S1
                ArrayEntity(11, count) = (R * L) / S2
            Else
                R = AlignSpiral.RadiusOut
                Dim S1 As Double = AlignSpiral.EndStation - PKStartAlign
                Dim S2 As Double = AlignSpiral.EndStation - PKEndAlign
                ArrayEntity(11, count) = (R * L) / S1
                ArrayEntity(10, count) = (R * L) / S2
            End If
        End If
        FuncCharacteristicsSpiral = True
    End Function
    'объект вложенная спираль
    Public Shared Function FuncCharacteristicsSubSpiral(ByVal align As Alignment, ByVal AlignSpiral As AlignmentSubEntitySpiral, ByRef ArrayEntity As Double(,), ByVal PKStartAlign As Double, ByVal PKEndAlign As Double) As Boolean
        FuncCharacteristicsSubSpiral = False
        If AlignSpiral.Length = 0 Then
            Exit Function
        End If
        Dim PKStartObj As Double = AlignSpiral.StartStation 'пикет начала дуги
        Dim PKEndObj As Double = AlignSpiral.EndStation 'пикет конца дуги
        Dim count As Integer = 0 'число элементов в массиве
        'определяем число элементов в массиве
        'условие до точки старта
        If PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj <= PKStartAlign And PKEndObj < PKEndAlign Then
            Exit Function
        End If
        If PKStartObj > PKStartAlign And PKStartObj >= PKEndAlign And PKEndObj > PKStartAlign And PKEndObj > PKEndAlign Then
            Exit Function
        End If
        If IsArray(ArrayEntity) = True Then
            count = ArrayEntity.GetUpperBound(1) + 1
        End If
        If (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignSpiral.StartStation 'пикет начала
            ArrayEntity(1, count) = AlignSpiral.EndStation 'пикет конца
            ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'координаты SPI точки
            ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
            ArrayEntity(4, count) = 2 'тип объекта
            ArrayEntity(5, count) = AlignSpiral.StartPoint.X 'координаты начала спирали
            ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
            ArrayEntity(7, count) = AlignSpiral.EndPoint.X 'координаты конца спирали
            ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
            ArrayEntity(9, count) = AlignSpiral.SPIStation 'пикет SPI точки
            Dim RIn As Double = AlignSpiral.RadiusIn
            If Double.IsInfinity(RIn) Then
                ArrayEntity(10, count) = -1 'входящий радиус
            Else
                ArrayEntity(10, count) = AlignSpiral.RadiusIn 'входящий радиус
            End If
            Dim ROut As Double = AlignSpiral.RadiusOut
            If Double.IsInfinity(ROut) Then
                ArrayEntity(11, count) = -1 'входящий радиус
            Else
                ArrayEntity(11, count) = AlignSpiral.RadiusOut 'входящий радиус
            End If
            ArrayEntity(12, count) = AlignSpiral.Length 'длина спирали
            ArrayEntity(14, count) = AlignSpiral.RadialPoint.X 'точка О спирали
            ArrayEntity(15, count) = AlignSpiral.RadialPoint.Y 'точка О спирали
            count += 1
            'объект выходящим концом находится внутри диапазона
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj <= PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign 'начало переходной кривой (ПК)
            ArrayEntity(1, count) = AlignSpiral.EndStation 'конец переходной кривой (ПК)
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim X As Double = -1 'начальные значения точки начала
            Dim Y As Double = -1
            Dim XY As Point2d = Nothing
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                XY = New Point2d(X, Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1 As Double = -1
            Dim Y1 As Double = -1
            Dim XY1 As Point2d = Nothing
            Dim PKPrev As Double = PKStartAlign - 0.001
            Try
                align.PointLocation(PKPrev, 0, X1, Y1)
                XY1 = New Point2d(X1, Y1)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2 As Double = -1
            Dim Y2 As Double = -1
            Dim XY2 As Point2d = Nothing
            Dim PKNext As Double = PKStartAlign + 0.001
            Try
                align.PointLocation(PKNext, 0, X2, Y2)
                XY2 = New Point2d(X2, Y2)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1 As Vector2d = Nothing
            Dim Vect2 As Vector2d = Nothing
            Dim Dir1 As Double = -1
            Dim Dir2 As Double = -1
            If X1 <> -1 And Y1 <> -1 And X <> -1 And Y <> -1 Then
                Vect1 = XY1.GetVectorTo(XY)
                Dir1 = Vect1.Angle
            End If
            If X2 <> -1 And Y2 <> -1 And X <> -1 And Y <> -1 Then
                Vect2 = XY.GetVectorTo(XY2)
                Dir2 = Vect2.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dir As Double = -1
            If Dir1 > -1 And Dir2 > -1 Then
                Dir = (Dir1 + Dir2) / 2
            ElseIf Dir > -1 And Dir2 = -1 Then
                Dir = Dir1
            ElseIf Dir1 = -1 And Dir2 > -1 Then
                Dir = Dir2
            End If
            'определяем новые координаты для направления первого вектора
            If Dir > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.SPIStation - AlignSpiral.StartStation)
                Dim Xv As Double = X + LenghtSPI * Math.Cos(Dir)
                Dim Yv As Double = Y + LenghtSPI * Math.Sin(Dir)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(X, Y, 0)
                Dim LineV1 As Line = New Line(XY13D, XYv3D)
                Dim XY23D As Point3d = New Point3d(AlignSpiral.EndPoint.X, AlignSpiral.EndPoint.Y, 0)
                Dim XYSPI3D As Point3d = New Point3d(AlignSpiral.SPIPoint.X, AlignSpiral.SPIPoint.Y, 0)
                Dim LineV2 As Line = New Line(XY23D, XYSPI3D)
                'ищем точку пересечения линий
                Dim ColPoint As Point3dCollection = New Point3dCollection
                LineV1.IntersectWith(LineV2, Intersect.ExtendBoth, New Plane(), ColPoint, IntPtr.Zero, IntPtr.Zero)
                If ColPoint.Count > 0 Then
                    For Each pt3d As Point3d In ColPoint
                        ArrayEntity(2, count) = pt3d.X 'новые координаты SPI точки для спирали
                        ArrayEntity(3, count) = pt3d.Y
                        ArrayEntity(5, count) = X
                        ArrayEntity(6, count) = Y
                        Exit For
                    Next
                End If
            Else
                ArrayEntity(0, count) = AlignSpiral.StartStation 'пикет начала сегмента
                ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'новые координаты SPI точки для спирали
                ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
                ArrayEntity(5, count) = AlignSpiral.StartPoint.X
                ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
            End If
            ArrayEntity(4, count) = 2
            ArrayEntity(7, count) = AlignSpiral.EndPoint.X
            ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
            ArrayEntity(9, count) = AlignSpiral.SPIStation
            'перезаписываем входящий радиус
            Dim S As Double = AlignSpiral.EndStation - PKStartAlign 'длина сегмента от начала конструкции до точки выбора
            Dim R As Double = AlignSpiral.RadiusOut
            'если все как положено входящий радиус бесконечен а выходящий имеет число
            If Double.IsInfinity(R) = False Then
                ArrayEntity(10, count) = (R * AlignSpiral.Length) / S
                ArrayEntity(11, count) = R
            Else
                'маловероятно но делаем все наоборот
                R = AlignSpiral.RadiusIn
                ArrayEntity(10, count) = R
                ArrayEntity(11, count) = (R * AlignSpiral.Length) / S
            End If
            ArrayEntity(12, count) = S 'длина объекта
            count += 1
            'объект входящим концом находится внутри диапазона
        ElseIf (PKStartObj >= PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = AlignSpiral.StartStation
            ArrayEntity(1, count) = PKEndAlign
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim X As Double = -1 'начальные значения точки конца
            Dim Y As Double = -1
            Dim XY As Point2d = Nothing
            Try
                align.PointLocation(PKEndAlign, 0, X, Y)
                XY = New Point2d(X, Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1 As Double = -1
            Dim Y1 As Double = -1
            Dim XY1 As Point2d = Nothing
            Dim PKPrev As Double = PKEndAlign - 0.001
            Try
                align.PointLocation(PKPrev, 0, X1, Y1)
                XY1 = New Point2d(X1, Y1)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2 As Double = -1
            Dim Y2 As Double = -1
            Dim XY2 As Point2d = Nothing
            Dim PKNext As Double = PKEndAlign + 0.001
            Try
                align.PointLocation(PKNext, 0, X2, Y2)
                XY2 = New Point2d(X2, Y2)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1 As Vector2d = Nothing
            Dim Vect2 As Vector2d = Nothing
            Dim Dir1 As Double = -1
            Dim Dir2 As Double = -1
            If X1 <> -1 And Y1 <> -1 And X <> -1 And Y <> -1 Then
                Vect1 = XY.GetVectorTo(XY1)
                Dir1 = Vect1.Angle
            End If
            If X2 <> -1 And Y2 <> -1 And X <> -1 And Y <> -1 Then
                Vect2 = XY2.GetVectorTo(XY)
                Dir2 = Vect2.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dir As Double = -1
            If Dir1 > -1 And Dir2 > -1 Then
                Dir = (Dir1 + Dir2) / 2
            ElseIf Dir > -1 And Dir2 = -1 Then
                Dir = Dir1
            ElseIf Dir1 = -1 And Dir2 > -1 Then
                Dir = Dir2
            End If
            'определяем новые координаты для направления первого вектора
            If Dir > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.EndStation - AlignSpiral.SPIStation)
                Dim Xv As Double = X + LenghtSPI * Math.Cos(Dir)
                Dim Yv As Double = Y + LenghtSPI * Math.Sin(Dir)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(X, Y, 0)
                Dim LineV1 As Line = New Line(XY13D, XYv3D)
                Dim XY23D As Point3d = New Point3d(AlignSpiral.StartPoint.X, AlignSpiral.StartPoint.Y, 0)
                Dim XYSPI3D As Point3d = New Point3d(AlignSpiral.SPIPoint.X, AlignSpiral.SPIPoint.Y, 0)
                Dim LineV2 As Line = New Line(XY23D, XYSPI3D)
                'ищем точку пересечения линий
                Dim ColPoint As Point3dCollection = New Point3dCollection
                LineV1.IntersectWith(LineV2, Intersect.ExtendBoth, New Plane(), ColPoint, IntPtr.Zero, IntPtr.Zero)
                If ColPoint.Count > 0 Then
                    For Each pt3d As Point3d In ColPoint
                        ArrayEntity(2, count) = pt3d.X 'новые координаты SPI точки для спирали
                        ArrayEntity(3, count) = pt3d.Y
                        ArrayEntity(7, count) = X
                        ArrayEntity(8, count) = Y
                        Exit For
                    Next
                End If
            Else
                ArrayEntity(1, count) = AlignSpiral.EndStation 'пикет начала сегмента
                ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'новые координаты SPI точки для спирали
                ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
                ArrayEntity(7, count) = AlignSpiral.EndPoint.X
                ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
            End If
            ArrayEntity(4, count) = 2
            ArrayEntity(5, count) = AlignSpiral.StartPoint.X
            ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
            ArrayEntity(9, count) = AlignSpiral.SPIStation
            'перезаписываем входящий радиус
            Dim S As Double = PKEndAlign - AlignSpiral.StartStation 'длина сегмента от начала конструкции до точки выбора
            Dim R As Double = AlignSpiral.RadiusIn
            'если все как положено входящий радиус бесконечен а выходящий имеет число
            If Double.IsInfinity(R) = False Then
                ArrayEntity(10, count) = R
                ArrayEntity(11, count) = (R * AlignSpiral.Length) / S
            Else
                'маловероятно но делаем все наоборот
                R = AlignSpiral.RadiusOut
                ArrayEntity(10, count) = (R * AlignSpiral.Length) / S
                ArrayEntity(11, count) = R
            End If
            ArrayEntity(12, count) = S 'длина объекта
            count += 1
            'объект обоими концами вне диапазона но середина объекта внутри
        ElseIf (PKStartObj < PKStartAlign And PKStartObj < PKEndAlign And PKEndObj > PKStartObj And PKEndObj > PKEndAlign) Then
            ReDim Preserve ArrayEntity(15, count)
            ArrayEntity(0, count) = PKStartAlign
            ArrayEntity(1, count) = PKEndAlign
            ArrayEntity(4, count) = 2
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim X As Double = -1 'начальные значения точки начала
            Dim Y As Double = -1
            Dim XY As Point2d = Nothing
            Try
                align.PointLocation(PKStartAlign, 0, X, Y)
                XY = New Point2d(X, Y)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1 As Double = -1
            Dim Y1 As Double = -1
            Dim XY1 As Point2d = Nothing
            Dim PKPrev As Double = PKStartAlign - 0.001
            Try
                align.PointLocation(PKPrev, 0, X1, Y1)
                XY1 = New Point2d(X1, Y1)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2 As Double = -1
            Dim Y2 As Double = -1
            Dim XY2 As Point2d = Nothing
            Dim PKNext As Double = PKStartAlign + 0.001
            Try
                align.PointLocation(PKNext, 0, X2, Y2)
                XY2 = New Point2d(X2, Y2)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1 As Vector2d = Nothing
            Dim Vect2 As Vector2d = Nothing
            Dim Dir1 As Double = -1
            Dim Dir2 As Double = -1
            If X1 <> -1 And Y1 <> -1 And X <> -1 And Y <> -1 Then
                Vect1 = XY1.GetVectorTo(XY)
                Dir1 = Vect1.Angle
            End If
            If X2 <> -1 And Y2 <> -1 And X <> -1 And Y <> -1 Then
                Vect2 = XY.GetVectorTo(XY2)
                Dir2 = Vect2.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dir As Double = -1
            If Dir1 > -1 And Dir2 > -1 Then
                Dir = (Dir1 + Dir2) / 2
            ElseIf Dir > -1 And Dir2 = -1 Then
                Dir = Dir1
            ElseIf Dir1 = -1 And Dir2 > -1 Then
                Dir = Dir2
            End If
            'определяем новые координаты для направления первого вектора
            Dim LineV1 As Line = New Line()
            If Dir > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.SPIStation - AlignSpiral.StartStation)
                Dim Xv As Double = X + LenghtSPI * Math.Cos(Dir)
                Dim Yv As Double = Y + LenghtSPI * Math.Sin(Dir)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(X, Y, 0)
                LineV1 = New Line(XY13D, XYv3D)
            End If
            '=========================================================================================
            'второй конец
            'находим координаты SPI точки (точки начала)
            '1. находим координаты начала
            Dim Xk As Double = -1 'начальные значения точки конца
            Dim Yk As Double = -1
            Dim XYk As Point2d = Nothing
            Try
                align.PointLocation(PKEndAlign, 0, Xk, Yk)
                XYk = New Point2d(Xk, Yk)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм назад
            Dim X1k As Double = -1
            Dim Y1k As Double = -1
            Dim XY1k As Point2d = Nothing
            Dim PKPrevk As Double = PKEndAlign - 0.001
            Try
                align.PointLocation(PKPrevk, 0, X1k, Y1k)
                XY1k = New Point2d(X1k, Y1k)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим координаты точки смещенной на 1 мм вперед
            Dim X2k As Double = -1
            Dim Y2k As Double = -1
            Dim XY2k As Point2d = Nothing
            Dim PKNextk As Double = PKEndAlign + 0.001
            Try
                align.PointLocation(PKNextk, 0, X2k, Y2k)
                XY2k = New Point2d(X2k, Y2k)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'находим 2 вектора
            Dim Vect1k As Vector2d = Nothing
            Dim Vect2k As Vector2d = Nothing
            Dim Dir1k As Double = -1
            Dim Dir2k As Double = -1
            If X1k <> -1 And Y1k <> -1 And Xk <> -1 And Yk <> -1 Then
                Vect1k = XYk.GetVectorTo(XY1k)
                Dir1k = Vect1k.Angle
            End If
            If X2k <> -1 And Y2k <> -1 And Xk <> -1 And Yk <> -1 Then
                Vect2k = XY2k.GetVectorTo(XYk)
                Dir2k = Vect2k.Angle
            End If
            'определяем среднее из дирекционных углов
            Dim Dirk As Double = -1
            If Dir1k > -1 And Dir2k > -1 Then
                Dirk = (Dir1k + Dir2k) / 2
            ElseIf Dirk > -1 And Dir2k = -1 Then
                Dirk = Dir1k
            ElseIf Dir1k = -1 And Dir2k > -1 Then
                Dirk = Dir2k
            End If
            'определяем новые координаты для направления первого вектора
            Dim LineV2 As Line = New Line()
            If Dirk > -1 Then
                Dim LenghtSPI As Double = Math.Abs(AlignSpiral.EndStation - AlignSpiral.SPIStation)
                Dim Xv As Double = Xk + LenghtSPI * Math.Cos(Dirk)
                Dim Yv As Double = Yk + LenghtSPI * Math.Sin(Dirk)
                Dim XYv3D As Point3d = New Point3d(Xv, Yv, 0)
                Dim XY13D As Point3d = New Point3d(Xk, Yk, 0)
                LineV2 = New Line(XY13D, XYv3D)
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ищем точки пересечения
            Dim ColPoint As Point3dCollection = New Point3dCollection
            LineV1.IntersectWith(LineV2, Intersect.ExtendBoth, New Plane(), ColPoint, IntPtr.Zero, IntPtr.Zero)
            If ColPoint.Count > 0 Then
                For Each pt3d As Point3d In ColPoint
                    ArrayEntity(2, count) = pt3d.X 'новые координаты SPI точки для спирали
                    ArrayEntity(3, count) = pt3d.Y
                    ArrayEntity(5, count) = X
                    ArrayEntity(6, count) = Y
                    ArrayEntity(7, count) = Xk
                    ArrayEntity(8, count) = Yk
                    ArrayEntity(12, count) = PKEndAlign - PKStartAlign 'длина объекта
                    Exit For
                Next
            Else
                ArrayEntity(2, count) = AlignSpiral.SPIPoint.X 'новые координаты SPI точки для спирали
                ArrayEntity(3, count) = AlignSpiral.SPIPoint.Y
                ArrayEntity(5, count) = AlignSpiral.StartPoint.X
                ArrayEntity(6, count) = AlignSpiral.StartPoint.Y
                ArrayEntity(7, count) = AlignSpiral.EndPoint.X
                ArrayEntity(8, count) = AlignSpiral.EndPoint.Y
                ArrayEntity(12, count) = AlignSpiral.EndStation - AlignSpiral.StartStation 'длина объекта
            End If
            'вычисляем радиусы
            Dim R As Double = AlignSpiral.RadiusIn
            Dim L As Double = AlignSpiral.Length
            If Double.IsInfinity(R) = False Then
                Dim S1 As Double = PKStartAlign - AlignSpiral.StartStation
                Dim S2 As Double = PKEndAlign - AlignSpiral.EndStation
                ArrayEntity(10, count) = (R * L) / S1
                ArrayEntity(11, count) = (R * L) / S2
            Else
                R = AlignSpiral.RadiusOut
                Dim S1 As Double = AlignSpiral.EndStation - PKStartAlign
                Dim S2 As Double = AlignSpiral.EndStation - PKEndAlign
                ArrayEntity(11, count) = (R * L) / S1
                ArrayEntity(10, count) = (R * L) / S2
            End If
        End If
        FuncCharacteristicsSubSpiral = True
    End Function

    '=============================================================================================================
    'функция проверяет сонаправленность входной спирали и дуги (конструкция СПИРАЛЬ ДУГА) (массив 0-1 координаты начала, 2-3 координаты конца, 4-5 SPI для спирали, 6 - Пикет точки SPI, 7-8 - координаты центра дуги или радиальная точка спирали)
    Public Shared Function FuncSpiralArcNapr1(ByVal align As Alignment, ByVal SP1 As Double(), ByVal Arc1 As Double()) As Boolean
        FuncSpiralArcNapr1 = False
        Try
            'считываем характеристики спирали
            Dim NSp1 As Point2d = New Point2d(SP1(0), SP1(1)) 'координаты начала переходной кривой
            Dim KSp1 As Point2d = New Point2d(SP1(2), SP1(3)) 'координаты конеца переходной кривой
            Dim SPISp1 As Point2d = New Point2d(SP1(4), SP1(5)) 'SPI точка переходной кривой
            'считываем характеристики дуги
            Dim NArc1 As Point2d = New Point2d(Arc1(0), Arc1(1)) 'координаты начала дуги
            Dim CArc1 As Point2d = New Point2d(Arc1(4), Arc1(5)) 'координаты центра дуги
            Dim L_SPI1_NSP1 As Vector2d = SPISp1.GetVectorTo(NSp1) 'вектор точка SPI - начало спирали
            Dim L_SPI1_NSP2 As Vector2d = SPISp1.GetVectorTo(NArc1) 'вектор точка SPI - начало дуги (конец спирали)
            'условие необходимо чтобы отсечь случаи принадлежности переходной кривой сразу двум смежным дугам или спиралям
            If L_SPI1_NSP2.Length < L_SPI1_NSP1.Length Then 'в случае если 2 вектор короче первого тогда первое условие пройдено
                '1. Самая первая проверка (если радиальная точка спирали совпадает с центром дуги спираль и дуга сонаправлены
                Dim Xarc As Double = Arc1(7) 'координата центра дуги X
                Dim Yarc As Double = Arc1(8) 'координата центра дуги Y
                Dim Xsp As Double = SP1(7) 'координата радиальной точки спирали X
                Dim Ysp As Double = SP1(8) 'координата радиальной точки спирали Y
                Dim Dx As Double = Math.Abs(Xarc - Xsp)
                Dim Dy As Double = Math.Abs(Yarc - Ysp)
                If Dx < 0.02 And Dy < 0.02 Then
                    'если координаты совпадают элементы сонаправлены
                    FuncSpiralArcNapr1 = True 'элементы сонаправлены
                    Exit Function
                End If
                '2 Проверка в случае неудачи попытаемся сделать еще проверку
                'вторая проверка на случай ели спираль выгнута и точка SPI лежит в той же плоскости что и центр дуги
                'проводим прямую линию от начала спирали до конца дуги
                Dim ax As Double = NSp1.X
                Dim ay As Double = NSp1.Y
                Dim bx As Double = Arc1(2)
                Dim by As Double = Arc1(3)
                Dim s1 As Double = (bx - ax) * (SPISp1.Y - ay) - (by - ay) * (SPISp1.X - ax)
                Dim s2 As Double = (bx - ax) * (CArc1.Y - ay) - (by - ay) * (CArc1.X - ax)
                If s1 > 0 And s2 < 0 Or s1 < 0 And s2 > 0 Then
                    FuncSpiralArcNapr1 = True 'спирали сонаправлены
                    Exit Function
                End If
                ax = NSp1.X
                ay = NSp1.Y
                align.PointLocation(Arc1(6), 0, bx, by)
                'пробуем сделать 3 проверку на случай если дуга развернутая
                s1 = (bx - ax) * (SPISp1.Y - ay) - (by - ay) * (SPISp1.X - ax)
                s2 = (bx - ax) * (CArc1.Y - ay) - (by - ay) * (CArc1.X - ax)
                If s1 > 0 And s2 < 0 Or s1 < 0 And s2 > 0 Then
                    FuncSpiralArcNapr1 = True 'спирали сонаправлены
                    Exit Function
                End If
            End If
        Catch
        End Try
    End Function
    '=============================================================================================================
    'функция проверяет сонаправленность выходной спирали и дуги (конструкция ДУГА СПИРАЛЬ) (массив 0-1 координаты начала, 2-3 координаты конца, 4-5 SPI для спирали, 6 - Пикет точки SPI, 7-8 - координаты центра дуги или радиальная точка спирали)
    Public Shared Function FuncArcSpiralNapr1(ByVal align As Alignment, ByVal Arc1 As Double(), ByVal SP1 As Double()) As Boolean
        FuncArcSpiralNapr1 = False
        Try
            'считываем характеристики дуги
            Dim KArc1 As Point2d = New Point2d(Arc1(2), Arc1(3)) 'координаты конца дуги
            Dim CArc1 As Point2d = New Point2d(Arc1(4), Arc1(5)) 'координаты центра дуги
            'считываем характеристики спирали
            Dim NSp1 As Point2d = New Point2d(SP1(0), SP1(1)) 'координаты начала переходной кривой
            Dim KSp1 As Point2d = New Point2d(SP1(2), SP1(3)) 'координаты конеца переходной кривой
            Dim SPISp1 As Point2d = New Point2d(SP1(4), SP1(5)) 'SPI точка переходной кривой
            Dim L_SPI1_KSP1 As Vector2d = SPISp1.GetVectorTo(KSp1) 'вектор точка SPI - конец спирали
            Dim L_SPI1_KArc As Vector2d = SPISp1.GetVectorTo(KArc1) 'вектор точка SPI - конец дуги (конец спирали)
            'условие необходимо чтобы отсечь случаи принадлежности переходной кривой сразу двум смежным дугам или спиралям
            If L_SPI1_KArc.Length < L_SPI1_KSP1.Length Then 'в случае если 2 вектор короче первого тогда первое условие пройдено
                '1. Самая первая проверка (если радиальная точка спирали совпадает с центром дуги спираль и дуга сонаправлены
                Dim Xarc As Double = Arc1(7) 'координата центра дуги X
                Dim Yarc As Double = Arc1(8) 'координата центра дуги Y
                Dim Xsp As Double = SP1(7) 'координата радиальной точки спирали X
                Dim Ysp As Double = SP1(8) 'координата радиальной точки спирали Y
                Dim Dx As Double = Math.Abs(Xarc - Xsp)
                Dim Dy As Double = Math.Abs(Yarc - Ysp)
                If Dx < 0.02 And Dy < 0.02 Then
                    'если координаты совпадают элементы сонаправлены
                    FuncArcSpiralNapr1 = True 'элементы сонаправлены
                    Exit Function
                End If
                'вторая проверка на случай ели спираль выгнута и точка SPI лежит в той же плоскости что и центр дуги
                Dim ax As Double = KSp1.X
                Dim ay As Double = KSp1.Y
                Dim bx As Double = Arc1(0)
                Dim by As Double = Arc1(1)
                Dim s1 As Double = (bx - ax) * (SPISp1.Y - ay) - (by - ay) * (SPISp1.X - ax)
                Dim s2 As Double = (bx - ax) * (CArc1.Y - ay) - (by - ay) * (CArc1.X - ax)
                If s1 > 0 And s2 < 0 Or s1 < 0 And s2 > 0 Then
                    FuncArcSpiralNapr1 = True 'спирали сонаправлены
                    Exit Function
                End If
                'третья проверка на случай если дуга развернута
                ax = KSp1.X
                ay = KSp1.Y
                align.PointLocation(Arc1(6), 0, bx, by)
                s1 = (bx - ax) * (SPISp1.Y - ay) - (by - ay) * (SPISp1.X - ax)
                s2 = (bx - ax) * (CArc1.Y - ay) - (by - ay) * (CArc1.X - ax)
                If s1 > 0 And s2 < 0 Or s1 < 0 And s2 > 0 Then
                    FuncArcSpiralNapr1 = True 'спирали сонаправлены
                    Exit Function
                End If
            End If
        Catch
        End Try
    End Function
    '=============================================================================================================
    'функция проверяет сонаправленность спиралей (конструкция СПИРАЛЬ СПИРАЛЬ)массив 0-1 координаты начала, 2-3 координаты конца, 4-5 SPI для спирали, 6 - Пикет точки SPI, 7-8 - координаты центра дуги или радиальная точка спирали
    Public Shared Function FuncSpiralSpiralNapr1(ByVal align As Alignment, ByVal SPin As Double(), ByVal SP2 As Double()) As Boolean
        FuncSpiralSpiralNapr1 = False
        Try
            Dim NSp1 As Point2d = New Point2d(SPin(0), SPin(1)) 'начало 1 переходной кривой
            Dim KSp1 As Point2d = New Point2d(SPin(2), SPin(3)) 'конец 1 переходной кривой
            Dim SPISp1 As Point2d = New Point2d(SPin(4), SPin(5)) 'SPI точка 1 переходной кривой
            Dim NSp2 As Point2d = New Point2d(SP2(0), SP2(1)) 'начало 1 переходной кривой
            Dim KSp2 As Point2d = New Point2d(SP2(2), SP2(3)) 'конец 1 переходной кривой
            Dim SPISp2 As Point2d = New Point2d(SP2(4), SP2(5)) 'SPI точка 1 переходной кривой
            'вычисляем длины векторов для певой спирали
            Dim L_SPI1_NSP1 As Vector2d = SPISp1.GetVectorTo(NSp1) 'точка SPIin первой спирали - начало 1 спирали
            Dim L_SPI1_NSP2 As Vector2d = SPISp1.GetVectorTo(NSp2) 'точка SPI первой спирали - начало второй спирали
            Dim Prov1 As Boolean = False
            'для перевой спирали если второй вектор короче первого первое подусловие пройдено
            If L_SPI1_NSP2.Length < L_SPI1_NSP1.Length Then Prov1 = True
            'вычисляем длины векторов для второй спирали
            Dim L_SPI2_KSP2 As Vector2d = SPISp2.GetVectorTo(KSp2)
            Dim L_SPI2_KSP1 As Vector2d = SPISp2.GetVectorTo(KSp1)
            Dim Prov2 As Boolean = False
            'для второй спирали если второй вектор короче первого первое подусловие пройдено
            If L_SPI2_KSP1.Length < L_SPI2_KSP2.Length Then Prov2 = True
            If Prov1 = True And Prov2 = True Then
                'условие необходимо чтобы отсечь случаи принадлежности переходной кривой сразу двум смежным дугам или спиралям
                '1. Самая первая проверка (если радиальная точка спирали совпадает с центром дуги спираль и дуга сонаправлены
                Dim Xarc As Double = SPin(7) 'координата центра дуги X
                Dim Yarc As Double = SPin(8) 'координата центра дуги Y
                Dim Xsp As Double = SP2(7) 'координата радиальной точки спирали X
                Dim Ysp As Double = SP2(8) 'координата радиальной точки спирали Y
                Dim Dx As Double = Math.Abs(Xarc - Xsp)
                Dim Dy As Double = Math.Abs(Yarc - Ysp)
                If Dx < 0.02 And Dy < 0.02 Then
                    'если координаты совпадают элементы сонаправлены
                    FuncSpiralSpiralNapr1 = True 'элементы сонаправлены
                    Exit Function
                End If
                'вторая проверка
                Dim ax As Double = NSp1.X
                Dim ay As Double = NSp1.Y
                Dim bx As Double = KSp2.X
                Dim by As Double = KSp2.Y
                Dim s1 As Double = (bx - ax) * (SPISp1.Y - ay) - (by - ay) * (SPISp1.X - ax)
                Dim s2 As Double = (bx - ax) * (SPISp2.Y - ay) - (by - ay) * (SPISp2.X - ax)
                If s1 > 0 And s2 > 0 Or s1 < 0 And s2 < 0 Then
                    FuncSpiralSpiralNapr1 = True 'спирали сонаправлены
                    Exit Function
                End If
            End If
        Catch
        End Try
    End Function
    '================================================================================================================================================================
    'функция разбивает кривые в плане с заданным шагом (трасса, Результирующий массив, Пикет начала, Пикет конца, Шаг разбивки, координаты вершины угла, массив с пикетажем кривой, коллекция рубленных пикетов, точность пикетов)
    Public Shared Function FuncSplitAlignCurvePlane(ByVal Align As Alignment, ByRef ArrayRezult As Double(,), ByVal StartRawPK As Double, ByVal EndRawPK As Double, ByRef ArrayK As Double(,), ByVal Shag As Double, ByVal VU As Point2d, ByVal SKPK As Double, Optional RoundZn As Integer = 2) As Boolean
        'ArrayK - массив с пикетажем кривых 0-НПК 1-НКК 2-ККК 3-КПК значения обязательны если нет переходных то НПК=НКК
        FuncSplitAlignCurvePlane = False
        Dim CountArrayRezult As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ищем середину дуги для расчета биссектрисы
        '1. Вычисляем координаты начала и конца конструкции
        'точка начала
        Dim XStart As Double = -1
        Dim YStart As Double = -1
        Try
            Align.PointLocation(StartRawPK, 0, XStart, YStart)
        Catch ex As Autodesk.Civil.PointNotOnEntityException
        End Try
        If XStart = -1 And YStart = -1 Then
            Try
                StartRawPK = StartRawPK + 0.01
                Align.PointLocation(StartRawPK, 0, XStart, YStart)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                Exit Function
            End Try
        End If
        Dim StartPoint As Point2d = New Point2d(XStart, YStart)
        'точка конца
        Dim XEnd As Double = -1
        Dim YEnd As Double = -1
        Try
            Align.PointLocation(EndRawPK, 0, XEnd, YEnd)
        Catch ex As Autodesk.Civil.PointNotOnEntityException
        End Try
        If XEnd = -1 And YEnd = -1 Then
            Try
                EndRawPK = EndRawPK - 0.01
                Align.PointLocation(EndRawPK, 0, XEnd, YEnd)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
                Exit Function
            End Try
        End If
        Dim EndPoint As Point2d = New Point2d(XEnd, YEnd)
        '=========================================================================================================
        '4. делаем вспомогательные линии
        Dim PTStart3d As Point3d = New Point3d(XStart, YStart, 0)
        Dim PTEnd3d As Point3d = New Point3d(XEnd, YEnd, 0)
        Dim PTVU3d As Point3d = New Point3d(VU.X, VU.Y, 0)
        Dim Line1 As Line = New Line(PTStart3d, PTVU3d)
        Dim Line2 As Line = New Line(PTVU3d, PTEnd3d)
        '==========================================================================================================
        '===========================================================================================================
        'заносим стартовый пикет в массив
        '============================================================================================================
        '============================================================================================================
        Dim ConstStartRawPiket As Double = StartRawPK
        ReDim Preserve ArrayRezult(8, CountArrayRezult)
        ArrayRezult(0, CountArrayRezult) = StartRawPK
        ArrayRezult(1, CountArrayRezult) = 0 'расстояние
        'вычисляем дельты
        ArrayRezult(2, CountArrayRezult) = 0 'dx
        ArrayRezult(3, CountArrayRezult) = 0 'dy
        ArrayRezult(4, CountArrayRezult) = XStart 'X
        ArrayRezult(5, CountArrayRezult) = YStart 'Y
        ArrayRezult(7, CountArrayRezult) = -1 'X
        ArrayRezult(8, CountArrayRezult) = -1 'Y
        'определяем описание (смотрим массив с информацией)
        Dim Desk As Double = 0
        If IsNothing(ArrayK) = False Then
            For i As Integer = 0 To ArrayK.GetUpperBound(1)
                Dim PKTemp As Double = ArrayK(0, i)
                Dim DeltaL As Double = Math.Abs(PKTemp - StartRawPK)
                If DeltaL < 0.01 Then
                    Desk = ArrayK(1, i)
                    ArrayK(2, i) = 1 'флаг что элемент записан
                    Exit For
                End If
            Next
        End If
        ArrayRezult(6, CountArrayRezult) = Desk
        CountArrayRezult += 1
        Desk = 0
        'смотрим рубленность для определения начального значения цикла
        Dim PKStartFormat As String = ""
        Try
            PKStartFormat = Align.GetStationStringWithEquations(StartRawPK) 'записываем пикет
        Catch ex As System.ArgumentException
            PKStartFormat = FuncFormatZN.FuncFormatPK(StartRawPK)
        End Try
        'проверям его на цельность
        Dim PKFormatDbl As Double = Val(PKStartFormat.Replace("+", ""))
        'вычленяем целую часть
        Dim IntPK = Math.Ceiling(PKFormatDbl / Shag) 'делим на шаг и округляем вниз
        IntPK = IntPK * Shag 'получаем целое значение кратное шагу
        If Math.Abs(IntPK - PKFormatDbl) > 0.0001 Then 'если значение не цельное тогда добавляем к начальному значению дельту
            Dim deltaPK As Double = IntPK - PKFormatDbl
            StartRawPK = StartRawPK + Math.Abs(deltaPK)
        Else
            StartRawPK = StartRawPK + Shag
        End If
        '=====================================================================================================================
        '5.Заполняем массив
        '=====================================================================================================================
        For i As Double = StartRawPK To EndRawPK Step Shag
            Dim PKFormat As String = "" 'форматированный пикет (начальное значение)
            '================================================================================
            Dim AlignXpt As Double = 0 'Координаты на трассе
            Dim AlignYPt As Double = 0
            Dim PtAlign2d As Point2d = Nothing
            Dim PtAlign3d As Point3d = Nothing
            Desk = 0
            '========================================================================================
            Dim pNearest1 As Point3d = Nothing 'координаты на линии
            '===================================================================================================
            'проверяем пикет на рубленность
            ''проверяем текущее значение на дробность
            Try
                PKFormat = Align.GetStationStringWithEquations(i) 'берем форматированный пикет
            Catch ex As System.ArgumentException
            End Try
            'проверям его на цельность
            PKFormatDbl = Val(PKFormat.Replace("+", "")) 'превращаем его в число
            'вычленяем целую часть
            IntPK = Math.Ceiling(PKFormatDbl / Shag) 'делим на шаг и округляем вниз
            IntPK = IntPK * Shag 'получаем целое значение кратное шагу
            If Math.Abs(IntPK - PKFormatDbl) > 0.001 Then 'если значение дробное то записываем его в массив
                Dim deltaPK As Double = IntPK - PKFormatDbl 'вычисляем шаг доворота
                i = i + Math.Abs(deltaPK)
            End If
            ''проверяем значение в массиве с дополнительными пикетами
            If IsNothing(ArrayK) = False Then 'если массив существует
                For j As Integer = 0 To ArrayK.GetUpperBound(1) 'проверяем циклом массив
                    Dim PKTemp As Double = ArrayK(0, j) 'пикет рубленный (RAW)
                    Dim DeltaLPK As Double = Math.Abs(i - PKTemp)
                    Dim flagPK As Integer = ArrayK(2, j)
                    If flagPK = 0 Then 'флаг равен нулю элемент нуждается в записи
                        If DeltaLPK < 0.01 Then 'если значение пикета совпало со значением доп.пикетов просто берем описание
                            Desk = ArrayK(1, j) 'смотрим описание
                            ArrayK(2, j) = 1
                        Else
                            If PKTemp < i Then 'если значение пикета больше значения доп пикета то пишем доп.пикет в массив
                                Try
                                    Align.PointLocation(PKTemp, 0, AlignXpt, AlignYPt)
                                Catch ex As Exception
                                End Try
                                PtAlign3d = New Point3d(AlignXpt, AlignYPt, 0)
                                ReDim Preserve ArrayRezult(8, CountArrayRezult)
                                ArrayRezult(0, CountArrayRezult) = PKTemp
                                If PKTemp <= SKPK Then
                                    pNearest1 = Line1.GetClosestPointTo(PtAlign3d, False)
                                    ArrayRezult(1, CountArrayRezult) = PKTemp - ConstStartRawPiket
                                    'вычисляем дельты
                                    Dim PtLin As Point2d = New Point2d(pNearest1.X, pNearest1.Y)
                                    PtAlign2d = New Point2d(AlignXpt, AlignYPt)
                                    ArrayRezult(2, CountArrayRezult) = StartPoint.GetVectorTo(PtLin).Length
                                    ArrayRezult(3, CountArrayRezult) = PtLin.GetVectorTo(PtAlign2d).Length
                                Else
                                    pNearest1 = Line2.GetClosestPointTo(PtAlign3d, False)
                                    ArrayRezult(1, CountArrayRezult) = EndRawPK - PKTemp
                                    'вычисляем дельты
                                    Dim PtLin As Point2d = New Point2d(pNearest1.X, pNearest1.Y)
                                    PtAlign2d = New Point2d(AlignXpt, AlignYPt)
                                    ArrayRezult(2, CountArrayRezult) = EndPoint.GetVectorTo(PtLin).Length
                                    ArrayRezult(3, CountArrayRezult) = PtLin.GetVectorTo(PtAlign2d).Length
                                End If
                                ArrayRezult(4, CountArrayRezult) = AlignXpt
                                ArrayRezult(5, CountArrayRezult) = AlignYPt
                                ArrayRezult(6, CountArrayRezult) = ArrayK(1, j)
                                ArrayRezult(7, CountArrayRezult) = pNearest1.X
                                ArrayRezult(8, CountArrayRezult) = pNearest1.Y
                                ArrayK(2, j) = 1
                                CountArrayRezult += 1
                            End If
                        End If
                    End If
                Next j
            End If
            ''================================================================================================
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'записываем основной пикет
            Try
                Align.PointLocation(i, 0, AlignXpt, AlignYPt)
            Catch ex As Exception
            End Try
            PtAlign3d = New Point3d(AlignXpt, AlignYPt, 0)
            ReDim Preserve ArrayRezult(8, CountArrayRezult)
            ArrayRezult(0, CountArrayRezult) = i
            If i <= SKPK Then
                pNearest1 = Line1.GetClosestPointTo(PtAlign3d, False)
                ArrayRezult(1, CountArrayRezult) = i - ConstStartRawPiket
                'вычисляем дельты
                Dim PtLin As Point2d = pNearest1.Convert2d(New Plane)
                PtAlign2d = New Point2d(AlignXpt, AlignYPt)
                ArrayRezult(2, CountArrayRezult) = StartPoint.GetVectorTo(PtLin).Length
                ArrayRezult(3, CountArrayRezult) = PtLin.GetVectorTo(PtAlign2d).Length
            Else
                pNearest1 = Line2.GetClosestPointTo(PtAlign3d, False)
                ArrayRezult(1, CountArrayRezult) = EndRawPK - i
                'вычисляем дельты
                Dim PtLin As Point2d = pNearest1.Convert2d(New Plane)
                PtAlign2d = New Point2d(AlignXpt, AlignYPt)
                ArrayRezult(2, CountArrayRezult) = EndPoint.GetVectorTo(PtLin).Length
                ArrayRezult(3, CountArrayRezult) = PtLin.GetVectorTo(PtAlign2d).Length
            End If
            ArrayRezult(4, CountArrayRezult) = AlignXpt
            ArrayRezult(5, CountArrayRezult) = AlignYPt
            ArrayRezult(6, CountArrayRezult) = Desk
            ArrayRezult(7, CountArrayRezult) = pNearest1.X
            ArrayRezult(8, CountArrayRezult) = pNearest1.Y
            CountArrayRezult += 1

        Next i
        '===========================================================================================
        'записываем конец конструкции
        Desk = 0
        Dim fldgEndPk As Boolean = False
        If IsArray(ArrayRezult) = True Then
            For i As Integer = 0 To ArrayRezult.GetUpperBound(1)
                Dim PK As Double = ArrayRezult(0, i)
                If PK = EndRawPK Then
                    fldgEndPk = True
                End If
            Next
        End If
        If fldgEndPk = False Then
            ReDim Preserve ArrayRezult(8, CountArrayRezult)
            ArrayRezult(0, CountArrayRezult) = EndRawPK
            ArrayRezult(1, CountArrayRezult) = 0 'расстояние
            'вычисляем дельты
            ArrayRezult(2, CountArrayRezult) = 0 'dx
            ArrayRezult(3, CountArrayRezult) = 0 'dy
            ArrayRezult(4, CountArrayRezult) = XEnd 'X
            ArrayRezult(5, CountArrayRezult) = YEnd 'Y
            ArrayRezult(7, CountArrayRezult) = -1 'X
            ArrayRezult(8, CountArrayRezult) = -1 'Y
            If IsNothing(ArrayK) = False Then
                For i As Integer = 0 To ArrayK.GetUpperBound(1)
                    Dim PKTemp As Double = ArrayK(0, i)
                    Dim DeltaL As Double = Math.Abs(PKTemp - EndRawPK)
                    If DeltaL < 0.01 Then
                        Desk = ArrayK(1, i)
                        ArrayK(2, i) = 1 'флаг что элемент нуждается в записи
                        Exit For
                    End If
                Next
            End If
            ArrayRezult(6, CountArrayRezult) = Desk
        End If
    End Function
    '================================================================================================================================================================
    'функция разбивает кривые в плане с заданным шагом (трасса, Результирующий массив, Пикет начала, Пикет конца, Шаг разбивки, координаты вершины угла, массив с пикетажем кривой, коллекция рубленных пикетов, точность пикетов)
    Public Shared Function FuncSplitAlignObjectStationXYHD(ByVal align As Alignment, ByVal Prof As Autodesk.Civil.DatabaseServices.Profile, ByRef ArrayObject(,) As Double, ByRef ArrayXYZD As Double(,), ByVal Shag As Double) As Boolean
        'ArrayK - массив с пикетажем кривых 0-НПК 1-НКК 2-ККК 3-КПК значения обязательны если нет переходных то НПК=НКК
        FuncSplitAlignObjectStationXYHD = False
        If IsNothing(align) = True Then Exit Function
        If IsArray(ArrayObject) = False Then
            Exit Function
        End If
        If align.Id.IsNull = True Then
            Exit Function
        End If
        If IsNothing(Prof) = True Then Exit Function
        If Prof.Id.IsNull = True Then
            Exit Function
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim CountArrayObject As Integer = 0
        Dim StartRawPiket As Double = ArrayObject(0, 0)
        Dim EndRawPiket As Double = ArrayObject(0, 1)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смотрим рубленность для определения начального значения цикла
        Dim PKStartFormat As String = ""
        Try
            PKStartFormat = align.GetStationStringWithEquations(StartRawPiket) 'записываем пикет
        Catch ex As System.ArgumentException
            PKStartFormat = FuncFormatZN.FuncFormatPK(StartRawPiket)
        End Try
        'проверям его на цельность
        Dim PKFormatDbl As Double = Val(PKStartFormat.Replace("+", ""))
        'вычленяем целую часть
        Dim IntPK = Math.Ceiling(PKFormatDbl / Shag) 'делим на шаг и округляем вниз
        IntPK = IntPK * Shag 'получаем целое значение кратное шагу
        If Math.Abs(IntPK - PKFormatDbl) > 0.0001 Then 'если значение не цельное тогда добавляем к начальному значению дельту
            Dim deltaPK As Double = IntPK - PKFormatDbl
            StartRawPiket = StartRawPiket + Math.Abs(deltaPK)
        Else
            StartRawPiket = StartRawPiket + Shag
        End If
        'определяем координаты XY
        Dim XStart As Double = -1
        Dim YStart As Double = -1
        Try
            align.PointLocation(ArrayObject(0, 0), 0, XStart, YStart)
        Catch ex As Autodesk.Civil.PointNotOnEntityException
        End Try
        'определяем высоту на профиле
        Dim HStart As Double = -1
        Try
            HStart = Prof.ElevationAt(ArrayObject(0, 0))
        Catch ex As System.ArgumentException
        End Try
        '=====================================================================================================================
        '5.Заполняем массив
        ReDim Preserve ArrayXYZD(4, CountArrayObject)
        ArrayXYZD(0, CountArrayObject) = ArrayObject(0, 0)
        ArrayXYZD(1, CountArrayObject) = XStart
        ArrayXYZD(2, CountArrayObject) = YStart
        ArrayXYZD(3, CountArrayObject) = HStart
        ArrayXYZD(4, CountArrayObject) = ArrayObject(1, 0)
        CountArrayObject += 1
        ArrayObject(2, 0) = 1
        '=====================================================================================================================
        For i As Double = StartRawPiket To EndRawPiket Step Shag
            Dim PKFormat As String = "" 'форматированный пикет (начальное значение)
            '================================================================================
            Dim AlignXpt As Double = -1 'Координаты на трассе
            Dim AlignYPt As Double = -1
            Dim ProfHPt As Double = -1
            Dim Desk As Integer = 0
            '===================================================================================================
            'проверяем пикет на рубленность
            Try
                PKFormat = align.GetStationStringWithEquations(i) 'берем форматированный пикет
                'проверям его на цельность
                PKFormatDbl = Val(PKFormat.Replace("+", "")) 'превращаем его в число
                'вычленяем целую часть
                IntPK = Math.Ceiling(PKFormatDbl / Shag) 'делим на шаг и округляем вниз
                IntPK = IntPK * Shag 'получаем целое значение кратное шагу
                If Math.Abs(IntPK - PKFormatDbl) > 0.001 Then 'если значение дробное то записываем его в массив
                    Dim deltaPK As Double = IntPK - PKFormatDbl 'вычисляем шаг доворота
                    i = i + Math.Abs(deltaPK)
                End If
            Catch ex As System.ArgumentException
            End Try
            ''проверяем значение в массиве с дополнительными пикетами
            If IsNothing(ArrayObject) = False Then 'если массив существует
                For j As Integer = 0 To ArrayObject.GetUpperBound(1) 'проверяем циклом массив
                    Dim PKTemp As Double = ArrayObject(0, j) 'пикет рубленный (RAW)
                    Dim DeltaLPK As Double = Math.Abs(i - PKTemp)
                    Dim flagPK As Integer = ArrayObject(2, j)
                    If flagPK = 0 Then 'флаг равен нулю элемент нуждается в записи
                        If DeltaLPK < 0.01 Then 'если значение пикета совпало со значением доп.пикетов просто берем описание
                            Desk = ArrayObject(1, j) 'смотрим описание
                        Else
                            If PKTemp < i Then 'если значение пикета больше значения доп пикета то пишем доп.пикет в массив
                                Try
                                    align.PointLocation(PKTemp, 0, AlignXpt, AlignYPt)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                                Try
                                    ProfHPt = Prof.ElevationAt(PKTemp)
                                Catch ex As System.ArgumentException
                                End Try
                                ReDim Preserve ArrayXYZD(4, CountArrayObject)
                                ArrayXYZD(0, CountArrayObject) = PKTemp
                                ArrayXYZD(1, CountArrayObject) = AlignXpt
                                ArrayXYZD(2, CountArrayObject) = AlignYPt
                                ArrayXYZD(3, CountArrayObject) = ProfHPt
                                ArrayXYZD(4, CountArrayObject) = ArrayObject(1, j)
                                ArrayObject(2, j) = 1
                                CountArrayObject += 1
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If
            ''================================================================================================
            ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'записываем основной пикет
            Try
                align.PointLocation(i, 0, AlignXpt, AlignYPt)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            Try
                ProfHPt = Prof.ElevationAt(i)
            Catch ex As System.ArgumentException
            End Try
            ReDim Preserve ArrayXYZD(4, CountArrayObject)
            ArrayXYZD(0, CountArrayObject) = i
            ArrayXYZD(1, CountArrayObject) = AlignXpt
            ArrayXYZD(2, CountArrayObject) = AlignYPt
            ArrayXYZD(3, CountArrayObject) = ProfHPt
            ArrayXYZD(4, CountArrayObject) = 0
            CountArrayObject += 1
        Next i
        '===========================================================================================
        'записываем конец конструкции
        Dim DeskEnd As Integer = 0
        Dim fldgEndPk As Boolean = False
        If IsArray(ArrayXYZD) = True Then
            For i As Integer = 0 To ArrayXYZD.GetUpperBound(1)
                Dim PK As Double = ArrayXYZD(0, i)
                If PK = EndRawPiket Then
                    fldgEndPk = True
                End If
            Next
        End If
        If fldgEndPk = False Then
            'определяем координаты XY
            Dim XEnd As Double = -1
            Dim YEnd As Double = -1
            Try
                align.PointLocation(EndRawPiket, 0, XEnd, YEnd)
            Catch ex As Autodesk.Civil.PointNotOnEntityException
            End Try
            'определяем высоту на профиле
            Dim HEnd As Double = -1
            Try
                HEnd = Prof.ElevationAt(ArrayObject(0, 1))
            Catch ex As System.ArgumentException
            End Try
            ReDim Preserve ArrayXYZD(4, CountArrayObject)
            ArrayXYZD(0, CountArrayObject) = EndRawPiket
            ArrayXYZD(1, CountArrayObject) = XEnd
            ArrayXYZD(2, CountArrayObject) = YEnd
            ArrayXYZD(3, CountArrayObject) = HEnd
            ArrayXYZD(4, CountArrayObject) = ArrayObject(1, 1)
        End If
        FuncSplitAlignObjectStationXYHD = True
    End Function

    '====================================================================================================================================================
    'функция заносит в массив пикеты конца и начала элементов вертикальной геометрии продольного профиля (0-Пикет, 1-наименование)
    Public Shared Function FuncReadVerticalGeometryToStatinName(ByVal ProfLine As Autodesk.Civil.DatabaseServices.Profile, ByRef ArrayVerticalGeomrtry As String(,), Optional ByVal RoundZn As Double = 3, Optional StartPK As Double = -1, Optional EndPK As Double = 9999999999) As Boolean
        FuncReadVerticalGeometryToStatinName = False
        If IsNothing(ProfLine) = True Then
            Exit Function
        End If
        If ProfLine.Id.IsNull = True Then
            Exit Function
        End If
        If ProfLine.Length = 0 Then
            Exit Function
        End If
        If StartPK = -1 Then
            StartPK = MathFunction.FuncTrimDigitFloo(ProfLine.StartingStation, RoundZn)
        End If
        If EndPK = 9999999999 Then
            EndPK = MathFunction.FuncTrimDigitCall(ProfLine.EndingStation, RoundZn)
        End If
        If StartPK > EndPK Then
            Dim TempZn As Double = StartPK
            StartPK = EndPK
            EndPK = TempZn
        End If
        Dim CountArrayProfPK As Integer = 0
        If IsNothing(ArrayVerticalGeomrtry) = False Then
            CountArrayProfPK = ArrayVerticalGeomrtry.GetUpperBound(1) + 1
        End If
        Dim EntCollection As ProfileEntityCollection = ProfLine.Entities
        'перебираем сегменты линии профиля
        Try
            For Each myAe As ProfileEntity In EntCollection
                Dim ProfEnt As ProfileEntityType = myAe.EntityType
                Dim st As Double = -1
                Dim en As Double = -1
                Dim NameStartSeg As String = "StartTangent"
                Dim NameEndSEg As String = "EndTangent"
                Select Case ProfEnt
                    'прямой участок
                    Case ProfileEntityType.Tangent
                        Dim TangentLine As ProfileTangent = myAe
                        If TangentLine.Length > 0 Then
                            st = TangentLine.StartStation
                            en = TangentLine.EndStation
                            NameStartSeg = "StartTangent"
                            NameEndSEg = "EndTangent"
                        End If
                    Case ProfileEntityType.Circular
                        Dim CircleLine As ProfileCircular = myAe
                        If CircleLine.Length > 0 Then
                            st = CircleLine.StartStation
                            en = CircleLine.EndStation
                            NameStartSeg = "StartCircular"
                            NameEndSEg = "EndCircular"
                        End If
                    Case ProfileEntityType.ParabolaAsymmetric
                        Dim ParabolaAsymLine As ProfileParabolaAsymmetric = myAe
                        If ParabolaAsymLine.Length > 0 Then
                            st = ParabolaAsymLine.StartStation
                            en = ParabolaAsymLine.EndStation
                            NameStartSeg = "StartParabolaAsymmetric"
                            NameEndSEg = "StartParabolaAsymmetric"
                        End If
                    Case ProfileEntityType.ParabolaSymmetric
                        Dim ParabolaSymLine As ProfileParabolaSymmetric = myAe
                        If ParabolaSymLine.Length > 0 Then
                            st = ParabolaSymLine.StartStation
                            en = ParabolaSymLine.EndStation
                            NameStartSeg = "StartParabolaSymmetric"
                            NameEndSEg = "EndParabolaSymmetric"
                        End If
                End Select
                If st <= StartPK And en <= StartPK Then
                    Continue For
                ElseIf st >= EndPK And en >= EndPK Then
                    Continue For
                ElseIf st <= StartPK And en >= EndPK Then
                    st = StartPK
                    en = EndPK
                ElseIf st < StartPK And en < endpk And en > startpk Then
                    st = StartPK
                ElseIf st < EndPK And st > startpk And en > EndPK Then
                    en = EndPK
                End If
                ReDim Preserve ArrayVerticalGeomrtry(1, CountArrayProfPK)
                ArrayVerticalGeomrtry(0, CountArrayProfPK) = st
                ArrayVerticalGeomrtry(1, CountArrayProfPK) = NameStartSeg
                CountArrayProfPK += 1
                ReDim Preserve ArrayVerticalGeomrtry(1, CountArrayProfPK)
                ArrayVerticalGeomrtry(0, CountArrayProfPK) = en
                ArrayVerticalGeomrtry(1, CountArrayProfPK) = NameEndSEg
                CountArrayProfPK += 1
            Next
            FuncReadVerticalGeometryToStatinName = True
        Catch ex As System.Exception
        End Try
    End Function

    '================================================================================================================================================================
    'функция разбивает элемент трассы с заданным шагом с учетом рубленности (результат массив 0-RawPiket 1-элемент 1-начало прямой 2-конец прямой 3-начало дуги 4-конец дуги 5-начало спирали 6-конец спирали 7-рубленный пикет)
    Public Shared Function FuncSplitAlignObjectRublPK(ByVal align As Alignment, ByRef ArrayObject() As Double, ByRef ArrayXYZD As Double(,), ByVal Shag As Double) As Boolean
        FuncSplitAlignObjectRublPK = False
        Dim count As Integer = 0
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'начало и конец элементов трассы
        Dim StartRawPiket As Double = ArrayObject(0)
        Dim EndRawPiket As Double = ArrayObject(1)
        Dim Kod As Integer = ArrayObject(2)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'рубленные пикеты
        'Dim stationEqw As StationEquationCollection = align.StationEquations
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'смотрим рубленность для определения начального значения цикла
        Dim PKStartFormat As String = ""
        Try
            PKStartFormat = align.GetStationStringWithEquations(StartRawPiket) 'записываем пикет
        Catch ex As System.ArgumentException
            PKStartFormat = FuncFormatZN.FuncFormatPK(StartRawPiket)
        End Try
        'проверям его на цельность
        Dim PKFormatDbl As Double = Val(PKStartFormat.Replace("+", ""))
        'вычленяем целую часть
        Dim IntPK = Math.Ceiling(PKFormatDbl / Shag) 'делим на шаг и округляем вниз
        IntPK = IntPK * Shag 'получаем целое значение кратное шагу
        If Math.Abs(IntPK - PKFormatDbl) > 0.0001 Then 'если значение не цельное тогда добавляем к начальному значению дельту
            Dim deltaPK As Double = IntPK - PKFormatDbl
            StartRawPiket = StartRawPiket + Math.Abs(deltaPK)
        Else
            StartRawPiket = StartRawPiket + Shag
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'записываем значение в массив
        ReDim Preserve ArrayXYZD(1, count)
        ArrayXYZD(0, count) = ArrayObject(0) 'записываем пикет начала элемента
        count += 1
        '=====================================================================================================================
        For i As Double = StartRawPiket To EndRawPiket Step Shag
            Dim PKFormat As String = "" 'форматированный пикет (начальное значение)
            '===================================================================================================
            'проверяем пикет на рубленность
            ''проверяем текущее значение на дробность
            Try
                PKFormat = align.GetStationStringWithEquations(i) 'берем форматированный пикет
            Catch ex As System.ArgumentException
                PKFormat = FuncFormatZN.FuncFormatPK(i)
            End Try
            'проверям его на цельность
            PKFormatDbl = Val(PKFormat.Replace("+", "")) 'превращаем его в число
            'вычленяем целую часть
            IntPK = Math.Ceiling(PKFormatDbl / Shag) 'делим на шаг и округляем вниз
            IntPK = IntPK * Shag 'получаем целое значение кратное шагу
            If Math.Abs(IntPK - PKFormatDbl) > 0.001 Then 'если значение дробное то записываем его в массив
                Dim deltaPK As Double = IntPK - PKFormatDbl 'вычисляем шаг доворота
                i = i + Math.Abs(deltaPK)
            End If
            ReDim Preserve ArrayXYZD(1, count)
            ArrayXYZD(0, count) = i
            count += 1
        Next i
        '===========================================================================================
        'записываем конец конструкции
        Dim DeskEnd As Integer = 0
        Dim fldgEndPk As Boolean = False
        If IsArray(ArrayXYZD) = True Then
            For i As Integer = 0 To ArrayXYZD.GetUpperBound(1)
                Dim PK As Double = ArrayXYZD(0, i)
                If PK = EndRawPiket Then
                    fldgEndPk = True
                End If
            Next
        End If
        If fldgEndPk = False Then
            ReDim Preserve ArrayXYZD(1, count)
            ArrayXYZD(0, count) = EndRawPiket
        End If
        '=========================================================================================
        'записываем коды элементов
        Dim indEnd As Integer = ArrayXYZD.GetUpperBound(1)
        If Kod = 0 Then
            ArrayXYZD(1, 0) = 1
            ArrayXYZD(1, indEnd) = 2
        ElseIf Kod = 1 Then
            ArrayXYZD(1, 0) = 3
            ArrayXYZD(1, indEnd) = 4
        ElseIf Kod = 2 Then
            ArrayXYZD(1, 0) = 5
            ArrayXYZD(1, indEnd) = 6
        End If
        FuncSplitAlignObjectRublPK = True
    End Function

    '=================================================================================================================================================================
    'функция получае массив с координатами и высотой с заданным шагом для всей трассы (горизонтальная геометрия не учитывается)
    Public Shared Function FuncReadAlignObjectXYZD(ByVal align As Autodesk.Civil.DatabaseServices.Alignment, ByVal Prof As Autodesk.Civil.DatabaseServices.Profile, ByVal ArrayObject() As Double, ByRef ArrayXYZD As String(,), ByVal Chag As Double, Optional ByVal ErrRound As Integer = 2, Optional ByRef sqwPKArray As Double(,) = Nothing) As Boolean
        'ArrayObject (0-пикет начала, 1-пикет конца, 2 - тип объекта (0-прямая, 1-дуга, 2-переходная кривая, 3-вся трасса)
        FuncReadAlignObjectXYZD = False
        Dim count As Integer = 0
        If IsNothing(align) = True Then Exit Function
        If IsNothing(Prof) = True Then Exit Function
        Dim StartPK As Double = ArrayObject(0) 'пикет начала
        Dim EndPK As Double = ArrayObject(1) 'пикет конца
        Dim TypeObject As Integer = ArrayObject(2) 'тип объекта
        '===============================================================================================================
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ветка для учета горизонтальной геометрии
        'разбиваем объект на части согласно значения шага
        Dim ChagRound As Double = Val(Chag)
        'получаем начальное значение цикла
        Dim IntPKStart As Integer = Math.Ceiling(StartPK / ChagRound) 'делим на шаг и округляем вверх начальный пикет
        IntPKStart = IntPKStart * ChagRound 'значение начала цикла
        'если значения не совпадают записываем значения стартового пикета в массив
        ReDim Preserve ArrayXYZD(4, count)
        'записываем значение пикета в массив
        Try
            Dim PKStartFormat As String = align.GetStationStringWithEquations(StartPK)
            ArrayXYZD(0, count) = PKStartFormat
        Catch ex As System.ArgumentException
            ArrayXYZD(0, count) = StartPK
        End Try
        'записываем координаты пикета
        Try
            Dim X As Double = 0
            Dim Y As Double = 0
            align.PointLocation(StartPK, 0, X, Y)
            X = Math.Round(X, ErrRound)
            Y = Math.Round(Y, ErrRound)
            ArrayXYZD(1, count) = X
            ArrayXYZD(2, count) = Y
        Catch ex As Autodesk.Civil.PointNotOnEntityException
        End Try
        'по пикетажу находим высоту на профиле
        Try
            Dim ElevProfStart As Double = Prof.ElevationAt(StartPK)
            ArrayXYZD(3, count) = Math.Round(ElevProfStart, ErrRound)
        Catch ex As System.ArgumentException
        End Try
        'записываем семантику
        ArrayXYZD(4, count) = "НТ"
        count += 1
        For i As Double = StartPK To EndPK Step Chag
            ReDim Preserve ArrayXYZD(4, count)
            Dim PKFormat As String = ""
            Dim X As Double = 0
            Dim Y As Double = 0
            Dim H As Double = 0
            '1. берем значения пикета
            Dim RawPiket As Double = i
            '2. получаем форматированный пикет
            Try
                PKFormat = align.GetStationStringWithEquations(RawPiket)
            Catch ex As System.ArgumentException
                PKFormat = RawPiket
            End Try
            '3. переводим его в число
            Dim PKFormatDbl As Double = Val(PKFormat.Replace("+", ""))
            '4. Проверяем его на дробность
            Dim IntPK = Math.Ceiling(PKFormatDbl / ChagRound) 'делим на шаг и округляем вниз
            IntPK = IntPK * ChagRound 'получаем целое значение кратное шагу
            If Math.Abs(IntPK - PKFormatDbl) > 0.0001 Then 'если значение цельное
                Dim deltaPK As Double = IntPK - PKFormatDbl
                i = i + Math.Abs(deltaPK)
                PKFormat = align.GetStationStringWithEquations(i) 'заново находим новое значение пикета
            End If
            '=================================================================================================
            'записываем основной пикет
            '=============================================================================================
            'проверяем новое значение пикета на соответсвие расстоянию больше 100м
            If IsArray(sqwPKArray) = True Then 'если массив с рубленными пикетами существует тогда обрабатываем его
                For j As Integer = 0 To sqwPKArray.GetUpperBound(1) 'перебираем массив с рубленными пикетами
                    Dim BaskPK As Double = sqwPKArray(1, j) 'значение предыдущего рубленного пикета
                    'вычленяем дробную часть
                    Dim IntBaskPk As Integer = Int(BaskPK / 100)
                    Dim DistPk As Double = (BaskPK / 100 - IntBaskPk) * 100
                    DistPk = Math.Round(DistPk, 6)
                    Dim Delta As Double = Math.Abs(PKFormatDbl - sqwPKArray(1, j))
                    Delta = Math.Round(Delta, 6)
                    If i < sqwPKArray(0, j) Then
                        If Delta <= DistPk And (sqwPKArray(2, j) - sqwPKArray(1, j)) < 0 Then
                            If sqwPKArray(3, j) = 0 Then
                                Dim PosPlus As Integer = PKFormat.IndexOf("+")
                                Dim pkInt As Integer = Left(PKFormat, PosPlus) - 1
                                Dim pkFloat As Double = Mid(PKFormat, PosPlus + 1) + 100
                                PKFormat = pkInt & "+" & FuncFormatZN.FuncFormatCoordinate(pkFloat, ErrRound)
                            End If
                        End If
                    End If
                Next j
            End If
            'находим координаты пикета
            Try
                align.PointLocation(i, 0, X, Y)
                X = Math.Round(X, ErrRound)
                Y = Math.Round(Y, ErrRound)
            Catch ex As System.ArgumentException
            End Try
            'по пикетажу находим высоту на профиле
            Try
                H = Prof.ElevationAt(i)
                H = Math.Round(H, ErrRound)
            Catch ex As System.ArgumentException
            End Try
            ArrayXYZD(0, count) = PKFormat
            ArrayXYZD(1, count) = X
            ArrayXYZD(2, count) = Y
            ArrayXYZD(3, count) = H
            ArrayXYZD(4, count) = "-"
            count += 1
            '6. вставляем в массив рубленный пикет
            Dim FlagRubPK As Boolean = False
            If IsArray(sqwPKArray) = True Then 'если массив с рубленными пикетами существует тогда обрабатываем его
                For j As Integer = 0 To sqwPKArray.GetUpperBound(1) 'перебираем массив с рубленными пикетами
                    'ищем ближайший к этому пикету рубленный пикет
                    Dim alPiket As Double = sqwPKArray(1, j) 'значение нового пикета
                    Dim DeltaRubPK As Double = alPiket - PKFormatDbl
                    'если расстояние меньше шага то ввинчиваем в массив рубленный пикет со всеми характеристиками
                    If DeltaRubPK <= Chag And sqwPKArray(3, j) = 0 Then
                        'записываем рубленный пикет после основного пикета
                        ReDim Preserve ArrayXYZD(4, count)
                        Dim RawRubPK As Double = sqwPKArray(0, j)
                        Try
                            ArrayXYZD(0, count) = align.GetStationStringWithEquations(RawRubPK)
                        Catch ex As System.ArgumentException
                            ArrayXYZD(0, count) = alPiket
                        End Try
                        'находим его координаты
                        Try
                            Dim RubX As Double = 0
                            Dim RubY As Double = 0
                            align.PointLocation(RawRubPK, 0, RubX, RubY)
                            ArrayXYZD(1, count) = Math.Round(RubX, ErrRound)
                            ArrayXYZD(2, count) = Math.Round(RubY, ErrRound)
                        Catch ex As System.ArgumentException
                        End Try
                        'находим высоту
                        Try
                            H = Prof.ElevationAt(RawRubPK)
                            ArrayXYZD(3, count) = Math.Round(H, ErrRound)
                        Catch ex As System.ArgumentException
                        End Try
                        ArrayXYZD(4, count) = "Руб.ПК"
                        sqwPKArray(3, j) = 1
                        count += 1
                        'корректируем значение i
                        'вычисляем добавку
                        i = RawRubPK - ChagRound
                        Exit For
                    End If
                Next j
            End If
line1:
        Next i
        '============================================================================================================
        'берем координаты конца сегмента
        Dim PKEndFormat As String = EndPK
        Try
            PKEndFormat = align.GetStationStringWithEquations(EndPK)
        Catch ex As System.ArgumentException
        End Try
        ReDim Preserve ArrayXYZD(4, count)
        'записываем значение пикета в массив
        ArrayXYZD(0, count) = PKEndFormat
        'записываем координаты пикета
        Try
            Dim X As Double = 0
            Dim Y As Double = 0
            align.PointLocation(EndPK, 0, X, Y)
            X = Math.Round(X, ErrRound)
            Y = Math.Round(Y, ErrRound)
            ArrayXYZD(1, count) = X
            ArrayXYZD(2, count) = Y
        Catch ex As Autodesk.Civil.PointNotOnEntityException
        End Try
        'по пикетажу находим высоту на профиле
        Try
            Dim ElevProfEnd As Double = Prof.ElevationAt(EndPK)
            ArrayXYZD(3, count) = Math.Round(ElevProfEnd, ErrRound)
        Catch ex As System.ArgumentException
        End Try
        'записываем семантику
        ArrayXYZD(4, count) = "КТ"
        FuncReadAlignObjectXYZD = True
    End Function
    '===============================================================================================================================
    '===============================================================================================================================
    'ФУНКЦИИ ДЛЯ РАБОТЫ С РУБЛЕННОСТЬЮ
    '===============================================================================================================================
    '===============================================================================================================================
    'функция определяет находится ли рубленный пикет в нужном диапазоне пикетов. Функция возвращает величину рубленности иначе 0. Все значения RAW
    Public Shared Function FuncFindRubkPK(ByVal align As Alignment, ByVal NK As Double, ByVal KK As Double) As Double
        FuncFindRubkPK = 0
        Dim ColPiketEq As StationEquationCollection = align.StationEquations
        If ColPiketEq.Count > 0 Then
            For Each PiketEq As StationEquation In ColPiketEq
                Dim RawSTBask As Double = PiketEq.RawStationBack
                If RawSTBask >= NK And RawSTBask < KK Then
                    Dim STBask As Double = PiketEq.StationBack
                    Dim STAhead As Double = PiketEq.StationAhead
                    'смотрим разницу 
                    Dim Delta_Bask_Ahead As Double = STAhead - STBask
                    FuncFindRubkPK = Delta_Bask_Ahead
                End If
            Next
        End If
    End Function
    'Функция форматирует пикет с учетом рубленности (трасса, исходный пикет, точность, формат вывода)0-пикет, 1 число, 2 километр
    Public Shared Function FuncFormatRublPiket(ByVal align As Alignment, ByVal RawPK As Double, Optional ByVal RoundZn As Integer = 2, Optional ByVal PKFormat As Integer = 0) As String
        FuncFormatRublPiket = ""
        'определяем рубленный пикет
        Dim PKStr As String = ""
        Try
            PKStr = align.GetStationStringWithEquations(RawPK)
            'проверяем на рубленность
            Dim sqwPKArray As StationEquationCollection = align.StationEquations
            If sqwPKArray.Count > 0 Then
                For i As Integer = 0 To sqwPKArray.Count - 1 'перебираем массив с рубленными пикетами
                    Dim RawStatBask As Double = sqwPKArray.Item(i).RawStationBack
                    Dim StatBask As Double = sqwPKArray.Item(i).StationBack
                    Dim StatAhead As Double = sqwPKArray.Item(i).StationAhead
                    'вычисляем дельту
                    Dim DeltaRubl As Double = StatAhead - StatBask
                    If DeltaRubl < 0 Then
                        'удлиненый пикет больше 100м
                        Dim PrewPK As Double = RawStatBask - Math.Abs(DeltaRubl) 'пикет RAW начала удлинения (положение рубленного пикета RAW минус дельта)
                        'получили RAW пикет начало удлиннения, конец удлиннения это положение рубленного пикета
                        'смотрим попал ли пикет в этот диапазон
                        If RawPK >= PrewPK And RawPK < RawStatBask Then
                            'пикет находится в диапазоне и подлежит порматированию с учетом более 100м
                            Dim PKSTRDbl As Double = Val(PKStr.Replace("+", ""))
                            'ищем целую часть
                            Dim IntPK As Integer = Math.Floor(PKSTRDbl / 100)
                            'ищем дробную часть
                            Dim FloafPK As Double = PKSTRDbl - IntPK * 100
                            'уменьшаем целую часть на единицу
                            Dim NewIntPK As Integer = IntPK - 1
                            If NewIntPK > -1 Then
                                FloafPK = FloafPK + 100
                            Else
                                NewIntPK = -1 * IntPK
                            End If
                            If PKFormat = 0 Then
                                Dim FLPKStr As String = FuncFormatZN.FuncFormatCoordinate(FloafPK, RoundZn)
                                FuncFormatRublPiket = IntPK & "+" & FLPKStr
                            ElseIf PKFormat = 1 Then
                                Dim FLPKStr As String = FuncFormatZN.FuncFormatCoordinate(FloafPK, RoundZn)
                                FuncFormatRublPiket = IntPK & FLPKStr
                            ElseIf PKFormat = 2 Then
                                Dim PKDbl As Double = IntPK * 100 + FloafPK
                                Dim IntKm As Integer = Math.Floor(PKDbl / 1000)
                                Dim FLKM As Double = (PKDbl / 1000 - IntKm) * 1000
                                Dim FLPKStr As String = FuncFormatZN.FuncFormatCoordinate(FLKM, RoundZn)
                                FuncFormatRublPiket = IntKm & "+" & FLPKStr
                            End If
                            Exit Function
                        End If
                    End If
                Next i
            End If
        Catch ex As System.ArgumentException
        End Try
        If PKFormat = 0 Then
            FuncFormatRublPiket = PKStr
        ElseIf PKFormat = 1 Then
            Dim PKDbl As Double = Val(PKStr.Replace("+", ""))
            FuncFormatRublPiket = FuncFormatZN.FuncFormatCoordinate(PKDbl, RoundZn)
        ElseIf PKFormat = 2 Then
            Dim PKDbl As Double = Val(PKStr.Replace("+", ""))
            Dim IntKm As Integer = Math.Floor(PKDbl / 1000)
            Dim FLKM As Double = (PKDbl / 1000 - IntKm) * 1000
            Dim FLPKStr As String = FuncFormatZN.FuncFormatCoordinate(FLKM, RoundZn)
            FuncFormatRublPiket = IntKm & "+" & FLPKStr
        End If
    End Function
    'вставка в чертеж рубленных пикетов
    Public Shared Function FuncInsertRublPK(ByVal align As Alignment, ByVal ArrayBLKProperties As String(,), Optional ByVal RoundZn As Integer = 2) As Integer
        FuncInsertRublPK = 0
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim RublPiketCollection As StationEquationCollection = align.StationEquations
        Using tr
            If RublPiketCollection.Count > 0 Then
                'проверяем есть ли в базе такой блок
                If IsArray(ArrayBLKProperties) = True Then
                    Dim acBlockTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    For i As Integer = 0 To ArrayBLKProperties.GetUpperBound(1)
                        Dim NameField As String = ArrayBLKProperties(0, i)
                        If IsNothing(NameField) = False Then
                            If NameField.Trim Like "Название блока" Then
                                Dim NameBlock As String = ArrayBLKProperties(1, i)
                                If IsNothing(NameBlock) = False Then
                                    If acBlockTbl.Has(NameBlock) = False Then
                                        Exit Function
                                    Else
                                        Exit For
                                    End If
                                End If
                            End If

                        End If
                    Next i
                End If
                'перебираем коллекцию
                For i As Integer = 0 To RublPiketCollection.Count - 1
                    Dim StRubl As StationEquation = RublPiketCollection.Item(i)
                    Dim RawBask As Double = StRubl.RawStationBack
                    Dim RawAhead As Double = StRubl.StationAhead
                    Dim StatBask As Double = StRubl.StationBack
                    Dim DeltaPK As Double = RawAhead - StatBask 'смотрим величину сдвига
                    Dim insBlk As Point3d = New Point3d(-1, -1, -1)
                    Dim LenghtBlk As Double = 100
                    Dim RotationBlk As Double = 0
                    'ищем координаты точек вставки
                    Dim IntPK As Integer = (Int(RawBask / 100)) * 100 'округляем значение пикета до целого 'Делаем пикет кратным 100
                    If DeltaPK < 0 Then 'пикет больше 100м
                        'сдвигаем его на 1 пикет вниз
                        Dim DPK1 As Double = StatBask / 100
                        Dim IntDPK1 As Integer = Math.Floor(StatBask) 'целая часть (ближайший целый пикет)
                        Dim NewDist As Double = IntDPK1 * 100 - 100 'длина рубленного пикета
                        'Dim NewDist As Double = StatBask - IntDPK1 * 100
                        NewDist = NewDist + 100 'получаем расстояние
                        Dim PKBastNew As Double = RawBask - NewDist ' рвсчетный raw-значение целого пикета
                        Dim AttrbLenghtTxt As String = FuncFormatZN.FuncFormatCoordinate(NewDist, RoundZn)
                        'ищем координаты точки вставки блока
                        Dim Xins As Double = -1
                        Dim Yins As Double = -1
                        Try
                            align.PointLocation(PKBastNew, 0, Xins, Yins)
                            insBlk = New Point3d(Xins, Yins, 0)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            GoTo line1
                        End Try
                        insBlk = New Point3d(Xins, Yins, 0)
                        'длина блока
                        'ищем угол доворота
                        'получаем координаты рубленного пикета
                        Dim XAh As Double = -1
                        Dim YAh As Double = -1
                        Dim InsBlk2 As Point3d = Nothing
                        Try
                            align.PointLocation(RawBask, 0, XAh, YAh)
                            Dim PtIns As Point2d = New Point2d(Xins, Yins)
                            Dim PtAh As Point2d = New Point2d(XAh, YAh)
                            Dim vect As Vector2d = PtIns.GetVectorTo(PtAh)
                            RotationBlk = vect.Angle
                            LenghtBlk = vect.Length
                            InsBlk2 = New Point3d(XAh, YAh, 0)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            RotationBlk = 0
                        End Try
                    Else
                        'пикет менее 100 м
                        Dim DPK1 As Double = StatBask / 100
                        Dim IntDPK1 As Integer = Math.Floor(DPK1)
                        Dim NewDist As Double = StatBask - IntDPK1 * 100
                        Dim AttrbLenghtTxt As String = FuncFormatZN.FuncFormatCoordinate(NewDist, RoundZn)
                        Dim PKBastNew As Double = RawBask - NewDist
                        'ищем координаты точки вставки блока
                        Dim Xins As Double = -1
                        Dim Yins As Double = -1
                        Try
                            align.PointLocation(PKBastNew, 0, Xins, Yins)
                            insBlk = New Point3d(Xins, Yins, 0)
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            GoTo line1
                        End Try
                        insBlk = New Point3d(Xins, Yins, 0)
                        'длина блока
                        'ищем угол доворота
                        'получаем координаты рубленного пикета
                        Dim XAh As Double = -1
                        Dim YAh As Double = -1
                        Dim InsBlk2 As Point3d = Nothing
                        Try
                            align.PointLocation(RawBask, 0, XAh, YAh)
                            Dim PtIns As Point2d = New Point2d(Xins, Yins)
                            Dim PtAh As Point2d = New Point2d(XAh, YAh)
                            Dim vect As Vector2d = PtIns.GetVectorTo(PtAh)
                            RotationBlk = vect.Angle
                            LenghtBlk = vect.Length
                        Catch ex As Autodesk.Civil.PointNotOnEntityException
                            RotationBlk = 0
                        End Try
                    End If
                    'вставляем блок
                    If insBlk.Z <> -1 Then
                        Dim RublBlk As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(insBlk, ArrayBLKProperties)
                        If IsNothing(RublBlk) = False Then
                            If IsDBNull(RublBlk) = False Then
                                FuncInsertRublPK += 1
                                '======================================================================================
                                'записываем атрибут
                                Dim GeoPoint As String(,) = Nothing
                                ReDim Preserve GeoPoint(1, 0)
                                GeoPoint(0, 0) = "DISTANCE"
                                GeoPoint(1, 0) = LenghtBlk
                                Dim boolWriteAttrib As Boolean = AcadFunction.FuncWriteAllBlockAttr(RublBlk, GeoPoint)
                                'меняем длину блока
                                If RublBlk.IsDynamicBlock = True Then
                                    Dim blockDyn As BlockTableRecord = tr.GetObject(RublBlk.DynamicBlockTableRecord, OpenMode.ForWrite)
                                    Dim col As DynamicBlockReferencePropertyCollection = RublBlk.DynamicBlockReferencePropertyCollection
                                    For Each dp As DynamicBlockReferenceProperty In col
                                        Dim values As Object() = dp.GetAllowedValues()
                                        Dim Name As String = dp.PropertyName
                                        If Name Like "Длина пикета" Then
                                            dp.Value = LenghtBlk
                                        ElseIf Name Like "Положение метки X" Then
                                            dp.Value = LenghtBlk / 2
                                        ElseIf Name Like "Положение метки Y" Then
                                            'dp.Value = Ytxt
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    End If
line1:
                Next i
            End If
            tr.Commit()
        End Using
    End Function

    '=================================================================================================================================
    '=================================================================================================================================
    'КОНСТРУКЦИИ
    '=================================================================================================================================
    '=================================================================================================================================
    'функция ищет ближайшую конструкцию к заданной точке (базовая линия, исходный пикет, начальное максимальное рассояние, допуск)
    Public Shared Function FuncFindBaseLineAssembly(ByVal bsLine As Baseline, ByVal StationPK As Double, ByVal ErrLine As Double, Optional ByVal dopusk As Double = 0) As AppliedAssembly
        FuncFindBaseLineAssembly = Nothing
        Dim IndAssembly As Integer = 0
        Dim CountAssembly As Integer = 0
        Do
            Try
                'получаем доступ к конструкции
                Dim appAssembly As AppliedAssembly = bsLine.GetAppliedAssemblyAtIndex(CountAssembly)
                Dim pointsAssebly As CalculatedPointCollection = appAssembly.Points
                For Each calkPoint As CalculatedPoint In pointsAssebly
                    Dim StationPT As Double = calkPoint.StationOffsetElevationToBaseline.X
                    Dim Dist As Double = Math.Abs(StationPK - StationPT)
                    If Dist < ErrLine Then
                        ErrLine = Dist
                        FuncFindBaseLineAssembly = appAssembly
                        Exit For
                    End If
                Next
                CountAssembly += 1
                If Not (ErrLine <= dopusk) Then
                    FuncFindBaseLineAssembly = Nothing
                End If
            Catch ex As System.ArgumentOutOfRangeException
                IndAssembly += 1
                If IndAssembly > 20 Then
                    Exit Do
                End If
            Catch ex As System.ArgumentException
                IndAssembly += 1
                If IndAssembly > 20 Then
                    Exit Do
                End If
            End Try
        Loop
    End Function


    '==================================================================================================================================
    'функция переводит значение пикетажа 0+00.00 в десятичное число 0.00000
    Public Shared Function FuncFormatPKtoDbl(ByVal PK As String) As Double
        FuncFormatPKtoDbl = 0
        Dim PKInt As Integer = 0
        Dim PKBbl As Double = 0
        Try
            'отделяем число полных пикетов
            Dim pos As Integer = PK.IndexOf("+")
            If pos > -1 Then
                'вычленяем целое число пикетов
                Dim PKIntStr As String = Mid(PK, 1, pos)
                PKIntStr = PKIntStr.Trim
                If IsNumeric(PKIntStr) = True Then
                    PKInt = Val(PKIntStr)
                End If
                Dim PKBblStr As String = Mid(PK, pos + 1)
                If IsNumeric(PKBblStr) = True Then
                    PKBbl = Val(PKBblStr)
                End If
            End If
        Catch
            Exit Function
        End Try
        FuncFormatPKtoDbl = PKInt * 100 + PKBbl
    End Function
End Class
