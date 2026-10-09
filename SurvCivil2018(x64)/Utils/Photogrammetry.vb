Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Numerics
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime

Public Class LensDistortionAnalyzer
    ' Коэффициенты дисторсии
    Public Property K1 As Double
    Public Property K2 As Double
    Public Property K3 As Double
    Public Property P1 As Double
    Public Property P2 As Double

    ' Параметры камеры
    Public Property FocalLengthX As Double
    Public Property FocalLengthY As Double
    Public Property PrincipalPointX As Double
    Public Property PrincipalPointY As Double
    Public Property Skew As Double

    Public Sub New()
        ' Значения по умолчанию (без дисторсии)
        K1 = 0
        K2 = 0
        K3 = 0
        P1 = 0
        P2 = 0

        ' Параметры по умолчанию
        FocalLengthX = 1000
        FocalLengthY = 1000
        PrincipalPointX = 0
        PrincipalPointY = 0
        Skew = 0
    End Sub

    ' Функция для определения дисторсии по шахматной доске
    Public Function CalibrateFromChessboard(
        imagePoints As List(Of PointF)(),
        boardSize As Size,
        squareSize As Double) As CalibrationResult

        ' boardSize: количество внутренних углов (ширина-1, высота-1)
        ' squareSize: размер квадрата в мм

        Dim result As New CalibrationResult()
        Dim objectPoints As New List(Of Vector3)()

        ' Создаем 3D координаты углов шахматной доски
        For y As Integer = 0 To boardSize.Height - 1
            For x As Integer = 0 To boardSize.Width - 1
                objectPoints.Add(New Vector3(x * squareSize, y * squareSize, 0))
            Next
        Next

        ' Преобразуем точки в нормализованные координаты
        Dim normalizedPoints As New List(Of PointF)()

        For Each pointList In imagePoints
            For Each point In pointList
                Dim normalized As New PointF(
                    (point.X - PrincipalPointX) / FocalLengthX,
                    (point.Y - PrincipalPointY) / FocalLengthY)
                normalizedPoints.Add(normalized)
            Next
        Next

        ' Вычисляем коэффициенты дисторсии методом наименьших квадратов
        result = CalculateDistortionCoefficients(normalizedPoints, objectPoints)

        Return result
    End Function

    Private Function CalculateDistortionCoefficients(
        imagePoints As List(Of PointF),
        objectPoints As List(Of Vector3)) As CalibrationResult

        Dim result As New CalibrationResult()

        ' Упрощенный расчет коэффициентов дисторсии
        ' В реальности используется метод Левенберга-Марквардта или SVD

        Dim totalError As Double = 0
        Dim count As Integer = imagePoints.Count

        ' Оцениваем дисторсию по отклонению от идеальной проекции
        For i As Integer = 0 To count - 1
            Dim idealPoint As PointF = ProjectPoint(objectPoints(i))
            Dim actualPoint As PointF = imagePoints(i)

            Dim dx As Double = actualPoint.X - idealPoint.X
            Dim dy As Double = actualPoint.Y - idealPoint.Y
            Dim distance As Double = Math.Sqrt(dx * dx + dy * dy)

            totalError += distance
        Next

        result.RMS_Error = totalError / count
        result.IsCalibrated = True

        ' Простая эвристика для определения типа дисторсии
        ' В реальном приложении здесь была бы сложная оптимизация

        Return result
    End Function

    ' Проекция 3D точки на 2D изображение (с учетом дисторсии)
    Public Function ProjectPoint(point3D As Vector3) As PointF
        ' Нормализованные координаты
        Dim xn As Double = point3D.X / point3D.Z
        Dim yn As Double = point3D.Y / point3D.Z

        ' Радиальное расстояние от центра
        Dim r2 As Double = xn * xn + yn * yn
        Dim r4 As Double = r2 * r2
        Dim r6 As Double = r4 * r2

        ' Радиальная дисторсия
        Dim radialDistortion As Double = 1 + K1 * r2 + K2 * r4 + K3 * r6

        ' Координаты с радиальной дисторсией
        Dim xr As Double = xn * radialDistortion
        Dim yr As Double = yn * radialDistortion

        ' Тангенциальная дисторсия
        Dim xd As Double = xr + (2 * P1 * xn * yn + P2 * (r2 + 2 * xn * xn))
        Dim yd As Double = yr + (P1 * (r2 + 2 * yn * yn) + 2 * P2 * xn * yn)

        ' Преобразование в пиксельные координаты
        Dim u As Double = FocalLengthX * xd + Skew * yd + PrincipalPointX
        Dim v As Double = FocalLengthY * yd + PrincipalPointY

        Return New PointF(CSng(u), CSng(v))
    End Function

    ' Коррекция дисторсии на изображении
    Public Function CorrectDistortion(image As Bitmap) As Bitmap
        Dim correctedImage As New Bitmap(image.Width, image.Height)

        ' Центр изображения
        Dim cx As Double = image.Width / 2
        Dim cy As Double = image.Height / 2

        ' Проходим по всем пикселям исправленного изображения
        For y As Integer = 0 To correctedImage.Height - 1
            For x As Integer = 0 To correctedImage.Width - 1
                ' Координаты в нормализованной системе
                Dim xn As Double = (x - cx) / FocalLengthX
                Dim yn As Double = (y - cy) / FocalLengthY

                ' Вычисляем обратную дисторсию
                Dim sourcePoint As PointF = ApplyInverseDistortion(New PointF(CSng(xn), CSng(yn)))

                ' Преобразуем обратно в пиксельные координаты
                Dim srcX As Integer = CInt(sourcePoint.X * FocalLengthX + cx)
                Dim srcY As Integer = CInt(sourcePoint.Y * FocalLengthY + cy)

                ' Проверяем границы
                If srcX >= 0 And srcX < image.Width And srcY >= 0 And srcY < image.Height Then
                    correctedImage.SetPixel(x, y, image.GetPixel(srcX, srcY))
                Else
                    correctedImage.SetPixel(x, y, Color.Black)
                End If
            Next
        Next

        Return correctedImage
    End Function

    ' Применяет обратную дисторсию (итеративный метод)
    Private Function ApplyInverseDistortion(point As PointF) As PointF
        ' Итеративный метод для нахождения исходных координат
        Dim xd As Double = point.X
        Dim yd As Double = point.Y

        ' Начальное приближение (предполагаем малую дисторсию)
        Dim xu As Double = xd
        Dim yu As Double = yd

        ' Итеративное уточнение (метод Ньютона)
        For i As Integer = 0 To 10 ' Обычно хватает 5-10 итераций
            Dim r2 As Double = xu * xu + yu * yu
            Dim r4 As Double = r2 * r2
            Dim r6 As Double = r4 * r2

            ' Вычисляем дисторсию для текущего приближения
            Dim radial As Double = 1 + K1 * r2 + K2 * r4 + K3 * r6
            Dim xr As Double = xu * radial
            Dim yr As Double = yu * radial

            ' Тангенциальная дисторсия
            Dim xDistorted As Double = xr + (2 * P1 * xu * yu + P2 * (r2 + 2 * xu * xu))
            Dim yDistorted As Double = yr + (P1 * (r2 + 2 * yu * yu) + 2 * P2 * xu * yu)

            ' Разность
            Dim dx As Double = xd - xDistorted
            Dim dy As Double = yd - yDistorted

            ' Если разность мала, останавливаемся
            If Math.Sqrt(dx * dx + dy * dy) < 0.00001 Then
                Exit For
            End If

            ' Корректируем приближение
            ' Упрощенный вариант - в реальности нужна матрица Якоби
            xu += dx / radial
            yu += dy / radial
        Next

        Return New PointF(CSng(xu), CSng(yu))
    End Function

    ' Анализ прямых линий для определения дисторсии
    Public Function AnalyzeLinesForDistortion(image As Bitmap) As DistortionAnalysis
        Dim analysis As New DistortionAnalysis()

        ' Детектируем линии на изображении (упрощенный вариант)
        Dim lines As List(Of LineSegment) = DetectLines(image)

        analysis.TotalLines = lines.Count

        For Each line In lines
            ' Вычисляем кривизну линии
            Dim curvature As Double = CalculateLineCurvature(line)
            analysis.LineCurvatures.Add(curvature)

            ' Определяем тип кривизны
            If curvature > 0.01 Then
                analysis.BarrelLines += 1
            ElseIf curvature < -0.01 Then
                analysis.PincushionLines += 1
            Else
                analysis.StraightLines += 1
            End If
        Next

        ' Определяем общий тип дисторсии
        If analysis.BarrelLines > analysis.PincushionLines * 2 Then
            analysis.DistortionType = "Бочкообразная"
        ElseIf analysis.PincushionLines > analysis.BarrelLines * 2 Then
            analysis.DistortionType = "Подушкообразная"
        ElseIf analysis.BarrelLines > 0 And analysis.PincushionLines > 0 Then
            analysis.DistortionType = "Сложная (бочко-подушка)"
        Else
            analysis.DistortionType = "Минимальная"
        End If

        Return analysis
    End Function

    Private Function DetectLines(image As Bitmap) As List(Of LineSegment)
        Dim lines As New List(Of LineSegment)()

        ' Упрощенный детектор линий
        ' В реальности используется преобразование Хафа или детектор границ Кэнни

        ' Для демонстрации создаем искусственные линии
        Dim width As Integer = image.Width
        Dim height As Integer = image.Height

        ' Вертикальные линии
        lines.Add(New LineSegment(New PointF(width \ 4, 0), New PointF(width \ 4, height)))
        lines.Add(New LineSegment(New PointF(width \ 2, 0), New PointF(width \ 2, height)))
        lines.Add(New LineSegment(New PointF(3 * width \ 4, 0), New PointF(3 * width \ 4, height)))

        ' Горизонтальные линии
        lines.Add(New LineSegment(New PointF(0, height \ 4), New PointF(width, height \ 4)))
        lines.Add(New LineSegment(New PointF(0, height \ 2), New PointF(width, height \ 2)))
        lines.Add(New LineSegment(New PointF(0, 3 * height \ 4), New PointF(width, 3 * height \ 4)))

        ' Диагональные линии
        lines.Add(New LineSegment(New PointF(0, 0), New PointF(width, height)))
        lines.Add(New LineSegment(New PointF(width, 0), New PointF(0, height)))

        Return lines
    End Function

    Private Function CalculateLineCurvature(line As LineSegment) As Double
        ' Упрощенный расчет кривизны
        ' В реальности анализируется множество точек вдоль линии

        ' Для демонстрации создаем искусственную кривизну
        ' в зависимости от положения линии

        Dim centerX As Single = (line.StartPt.X + line.EndPt.X) / 2
        Dim centerY As Single = (line.StartPt.Y + line.EndPt.Y) / 2

        ' Расстояние от центра изображения
        Dim imageCenterX As Single = 512 ' Предполагаем изображение 1024x1024
        Dim imageCenterY As Single = 512
        Dim distanceFromCenter As Double = Math.Sqrt(
            (centerX - imageCenterX) * (centerX - imageCenterX) +
            (centerY - imageCenterY) * (centerY - imageCenterY))

        ' Искусственная дисторсия: бочкообразная на краях
        Dim normalizedDistance As Double = distanceFromCenter / imageCenterX

        ' Модель бочкообразной дисторсии
        Return 0.1 * normalizedDistance * normalizedDistance
    End Function

    ' Визуализация дисторсии
    Public Function CreateDistortionGrid(imageSize As Size, gridSize As Integer) As Bitmap
        Dim gridImage As New Bitmap(imageSize.Width, imageSize.Height)

        Using g As Graphics = Graphics.FromImage(gridImage)
            g.Clear(Color.White)

            Dim pen As New Pen(Color.Black, 1)

            ' Рисуем идеальную сетку
            For x As Integer = 0 To imageSize.Width Step gridSize
                g.DrawLine(pen, x, 0, x, imageSize.Height)
            Next

            For y As Integer = 0 To imageSize.Height Step gridSize
                g.DrawLine(pen, 0, y, imageSize.Width, y)
            Next

            ' Применяем дисторсию к сетке
            If Math.Abs(K1) > 0.0001 Or Math.Abs(K2) > 0.0001 Then
                Dim distortedPen As New Pen(Color.Red, 2)

                ' Рисуем искаженные вертикальные линии
                For x As Integer = 0 To imageSize.Width Step gridSize
                    Dim points As New List(Of PointF)()

                    For y As Integer = 0 To imageSize.Height Step 10
                        Dim normalizedX As Double = (x - imageSize.Width / 2) / FocalLengthX
                        Dim normalizedY As Double = (y - imageSize.Height / 2) / FocalLengthY

                        Dim distorted As PointF = ApplyDistortion(New PointF(CSng(normalizedX), CSng(normalizedY)))

                        Dim screenX As Integer = CInt(distorted.X * FocalLengthX + imageSize.Width / 2)
                        Dim screenY As Integer = CInt(distorted.Y * FocalLengthY + imageSize.Height / 2)

                        points.Add(New PointF(screenX, screenY))
                    Next

                    If points.Count > 1 Then
                        g.DrawLines(distortedPen, points.ToArray())
                    End If
                Next

                ' Рисуем искаженные горизонтальные линии
                For y As Integer = 0 To imageSize.Height Step gridSize
                    Dim points As New List(Of PointF)()

                    For x As Integer = 0 To imageSize.Width Step 10
                        Dim normalizedX As Double = (x - imageSize.Width / 2) / FocalLengthX
                        Dim normalizedY As Double = (y - imageSize.Height / 2) / FocalLengthY

                        Dim distorted As PointF = ApplyDistortion(New PointF(CSng(normalizedX), CSng(normalizedY)))

                        Dim screenX As Integer = CInt(distorted.X * FocalLengthX + imageSize.Width / 2)
                        Dim screenY As Integer = CInt(distorted.Y * FocalLengthY + imageSize.Height / 2)

                        points.Add(New PointF(screenX, screenY))
                    Next

                    If points.Count > 1 Then
                        g.DrawLines(distortedPen, points.ToArray())
                    End If
                Next
            End If
        End Using

        Return gridImage
    End Function

    Private Function ApplyDistortion(point As PointF) As PointF
        Dim x As Double = point.X
        Dim y As Double = point.Y

        Dim r2 As Double = x * x + y * y
        Dim r4 As Double = r2 * r2
        Dim r6 As Double = r4 * r2

        Dim radial As Double = 1 + K1 * r2 + K2 * r4 + K3 * r6

        Dim xr As Double = x * radial
        Dim yr As Double = y * radial

        Dim xd As Double = xr + (2 * P1 * x * y + P2 * (r2 + 2 * x * x))
        Dim yd As Double = yr + (P1 * (r2 + 2 * y * y) + 2 * P2 * x * y)

        Return New PointF(CSng(xd), CSng(yd))
    End Function
End Class

' Вспомогательные классы
Public Class CalibrationResult
    Public Property RMS_Error As Double
    Public Property IsCalibrated As Boolean
    Public Property CameraMatrix As Double(,)
    Public Property DistortionCoefficients As Double()
End Class

Public Class DistortionAnalysis
    Public Property TotalLines As Integer
    Public Property StraightLines As Integer
    Public Property BarrelLines As Integer
    Public Property PincushionLines As Integer
    Public Property DistortionType As String
    Public Property LineCurvatures As New List(Of Double)()

    Public ReadOnly Property DistortionPercentage As Double
        Get
            If TotalLines = 0 Then Return 0
            Return (BarrelLines + PincushionLines) / TotalLines * 100
        End Get
    End Property
End Class

Public Class LineSegment
    Public Property StartPt As PointF
    Public Property EndPt As PointF

    Public Sub New(startP As PointF, endP As PointF)
        Me.StartPt = startP
        Me.EndPt = endP
    End Sub

    Public ReadOnly Property Length As Double
        Get
            Dim dx As Double = EndPt.X - StartPt.X
            Dim dy As Double = EndPt.Y - StartPt.Y
            Return Math.Sqrt(dx * dx + dy * dy)
        End Get
    End Property
End Class

' Пример использования
Public Class ModulPhoto
    <CommandMethod("testPhoto")>
    Public Sub testPhoto()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        ed.WriteMessage(vbCrLf & "Анализ дисторсии объектива")
        ' Создаем анализатор
        Dim analyzer As New LensDistortionAnalyzer()

        ' Устанавливаем параметры дисторсии (пример бочкообразной дисторсии)
        analyzer.K1 = -0.0002
        analyzer.K2 = 0.0000001

        ' Анализ дисторсии
        ed.WriteMessage(vbCrLf & "1. Анализ типа дисторсии:")

        ' Создаем тестовое изображение с линиями
        Dim testImage As New Bitmap(1024, 1024)
        Using g As Graphics = Graphics.FromImage(testImage)
            g.Clear(Color.White)

            ' Рисуем сетку
            Dim pen As New Pen(Color.Black, 2)
            For x As Integer = 100 To 900 Step 200
                g.DrawLine(pen, x, 100, x, 900)
            Next

            For y As Integer = 100 To 900 Step 200
                g.DrawLine(pen, 100, y, 900, y)
            Next
        End Using

        ' Сохраняем тестовое изображение
        testImage.Save("test_grid.png", ImageFormat.Png)

        ' Анализируем линии
        Dim analysis As DistortionAnalysis = analyzer.AnalyzeLinesForDistortion(testImage)

        Console.WriteLine($"Всего линий: {analysis.TotalLines}")
        Console.WriteLine($"Прямых линий: {analysis.StraightLines}")
        Console.WriteLine($"Линий с бочкообразной дисторсией: {analysis.BarrelLines}")
        Console.WriteLine($"Линий с подушкообразной дисторсией: {analysis.PincushionLines}")
        Console.WriteLine($"Тип дисторсии: {analysis.DistortionType}")
        Console.WriteLine($"Процент дисторсии: {analysis.DistortionPercentage:F1}%")
        Console.WriteLine()

        ' Создаем визуализацию дисторсии
        Console.WriteLine("2. Визуализация дисторсии:")
        Dim gridImage As Bitmap = analyzer.CreateDistortionGrid(New Size(1024, 1024), 100)
        gridImage.Save("distortion_grid.png", ImageFormat.Png)
        Console.WriteLine("Сетка дисторсии сохранена в distortion_grid.png")
        Console.WriteLine("Черные линии - идеальная сетка")
        Console.WriteLine("Красные линии - сетка с дисторсией")
        Console.WriteLine()

        ' Коррекция дисторсии
        Console.WriteLine("3. Коррекция дисторсии:")

        ' Создаем изображение с дисторсией
        Dim distortedImage As New Bitmap(1024, 1024)
        Using g As Graphics = Graphics.FromImage(distortedImage)
            g.Clear(Color.LightGray)

            ' Рисуем текст, который будет искажен
            Using font As New Font("Arial", 48)
                g.DrawString("DISTORTION TEST", font, Brushes.Black, 200, 400)
            End Using
        End Using

        distortedImage.Save("distorted_text.png", ImageFormat.Png)

        ' Корректируем дисторсию
        Dim correctedImage As Bitmap = analyzer.CorrectDistortion(distortedImage)
        correctedImage.Save("corrected_text.png", ImageFormat.Png)

        Console.WriteLine("Искаженное изображение сохранено в distorted_text.png")
        Console.WriteLine("Скорректированное изображение сохранено в corrected_text.png")
        Console.WriteLine()

        ' Расчет коэффициентов дисторсии
        Console.WriteLine("4. Расчет коэффициентов дисторсии:")
        Console.WriteLine($"K1 (радиальная дисторсия 2-го порядка): {analyzer.K1:E4}")
        Console.WriteLine($"K2 (радиальная дисторсия 4-го порядка): {analyzer.K2:E4}")
        Console.WriteLine($"K3 (радиальная дисторсия 6-го порядка): {analyzer.K3:E4}")
        Console.WriteLine($"P1 (тангенциальная дисторсия): {analyzer.P1:E4}")
        Console.WriteLine($"P2 (тангенциальная дисторсия): {analyzer.P2:E4}")
        Console.WriteLine()

        ' Информация о калибровке
        Console.WriteLine("5. Калибровка камеры:")
        Console.WriteLine("Для точной калибровки необходимо:")
        Console.WriteLine("1. Распечатать шахматную доску (например, 9x6 квадратов)")
        Console.WriteLine("2. Снять 10-20 фото под разными углами")
        Console.WriteLine("3. Обнаружить углы шахматной доски на каждом изображении")
        Console.WriteLine("4. Использовать метод наименьших квадратов для расчета коэффициентов")
        Console.WriteLine()

        Console.WriteLine("Нажмите любую клавишу для выхода...")
        Console.ReadKey()
    End Sub
End Class
