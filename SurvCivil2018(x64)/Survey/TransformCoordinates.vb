Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Public Class TransformCoordinates
    Inherits System.Windows.Forms.Form
    Public stbool As Boolean = False
    Public Sub New()
        InitializeComponent()
    End Sub
    'НАЖАТА КНОПКА ОТМЕНА
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        stbool = False
        Me.Hide()
    End Sub
    'НАЖАТА КНОПКА ЗАГРУЗИТЬ ИСХОДНУЮ СИСТЕМУ КООРДИНАТ
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim RowCount As Integer = 0
        Dim count As Integer = 0
        Dim FileNameGeo As String = ""
        OpenFileDialog1.Filter = "Все|*.*"
        OpenFileDialog1.Title = "Выбор координатного файла"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FileNameGeo = OpenFileDialog1.FileName
        End If
        Try
            Dim input As StreamReader = New StreamReader(FileNameGeo, True)
            Dim line1 As String
            Do Until input.EndOfStream
                line1 = input.ReadLine() 'считываем строку
                line1 = line1.Trim
                If line1.Trim.Length = 0 Then Exit Do
                Dim Blok As String() = line1.Split
                DataGridView1.Rows.Insert(RowCount)
                For i As Integer = 0 To Blok.GetUpperBound(0)
                    Dim Rez As String = Blok(i)
                    If Rez.Trim.Length > 0 Then
                        count += 1
                        DataGridView1.Rows(RowCount).Cells(count).Value = Blok(i)
                    End If
                    If count > 2 Then Exit For
                Next
                DataGridView1.Rows(RowCount).Cells(0).Value = True
                RowCount += 1
                count = 0
            Loop
        Catch
            MsgBox("Не удалось загрузить файл. Влзможно в сроке " & RowCount & " содержится ошибка")
        End Try
    End Sub
    'НАЖАТА КНОПКА ЗАГРУЗИТЬ МЕСТНУЮ СИСТЕМУ КООРДИНАТ
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim RowCount As Integer = 0
        Dim count As Integer = 0
        Dim FileNameGeo As String = ""
        OpenFileDialog1.Filter = "Все|*.*"
        OpenFileDialog1.Title = "Выбор координатного файла"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            FileNameGeo = OpenFileDialog1.FileName
        End If
        Try
            Dim input As StreamReader = New StreamReader(FileNameGeo, True)
            Dim line1 As String
            Do Until input.EndOfStream
                line1 = input.ReadLine() 'считываем строку
                line1 = line1.Trim
                If line1.Trim.Length = 0 Then Exit Do
                Dim Blok As String() = line1.Split
                DataGridView2.Rows.Insert(RowCount)
                For i As Integer = 0 To Blok.GetUpperBound(0)
                    Dim Rez As String = Blok(i)
                    If Rez.Trim.Length > 0 Then
                        count += 1
                        DataGridView2.Rows(RowCount).Cells(count).Value = Blok(i)
                    End If
                    If count > 2 Then Exit For
                Next
                DataGridView2.Rows(RowCount).Cells(0).Value = True
                RowCount += 1
                count = 0
            Loop
        Catch
            MsgBox("Не удалось загрузить файл. Влзможно в сроке " & RowCount & " содержится ошибка")
        End Try
    End Sub
    'НАЖАТА КНОПКА РАССЧИТАТЬ ПАРАМЕТРЫ
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '1.исходные данные
        Dim ScaleFactor As Double = 0 'масштабный фактор
        Dim RotationAngle As Double = 0 'угол поворота в радианах
        Dim DX As Double = 0 'смещение по оси Х в метрах
        Dim DY As Double = 0 'смещение по оси Y в метрах
        Dim ArrayMSK As String(,) = Nothing 'массив для хранения координат местной МСК
        Dim ArraySK As String(,) = Nothing 'массив для хранения координат глобальной СК
        '2.массивы для хранения расчетных данных
        Dim ArrayCount As Integer = 0
        Dim ArrayDeltaX As Double() = Nothing
        Dim ArrayDeltaY As Double() = Nothing
        '3.записываем в массив координатные данные
        Dim RowTable1 As Integer = DataGridView1.Rows.Count - 1 'количество записей в таблице 1 (СК)
        Dim RowTable2 As Integer = DataGridView2.Rows.Count - 1 'количество записей в таблице 2 (МСК)
        'проверяем чтобы количество записей было одинаково в обеих таблицах
        Dim count As Integer = 0
        If RowTable1 < RowTable2 Then
            count = RowTable1
        Else
            count = RowTable2
        End If
        For i As Integer = 0 To count - 1
            If DataGridView1.Rows(i).Cells(0).Value = True Then
                ReDim Preserve ArraySK(2, ArrayCount)
                ArraySK(0, ArrayCount) = DataGridView1.Rows(i).Cells(3).Value 'записываем Y
                ArraySK(1, ArrayCount) = DataGridView1.Rows(i).Cells(2).Value 'записываем X 
                ArraySK(2, ArrayCount) = DataGridView1.Rows(i).Cells(1).Value 'записываем номер 
                ReDim Preserve ArrayMSK(2, i)
                ArrayMSK(0, ArrayCount) = DataGridView2.Rows(i).Cells(3).Value 'записываем Y
                ArrayMSK(1, ArrayCount) = DataGridView2.Rows(i).Cells(2).Value 'записываем X 
                ArrayMSK(2, ArrayCount) = DataGridView2.Rows(i).Cells(1).Value 'записываем номер 
                ArrayCount += 1
            End If
        Next i
        '4.находим масштабный фактор
        Dim CountScaleFactor As Double = 0
        For i As Integer = 0 To ArraySK.GetUpperBound(1) - 1
            For j As Integer = i + 1 To ArraySK.GetUpperBound(1)
                'СК
                Dim pt3 As Point2d = New Point2d(ArraySK(0, i), ArraySK(1, i))
                Dim pt4 As Point2d = New Point2d(ArraySK(0, j), ArraySK(1, j))
                Dim lenghtSK As Double = pt3.GetVectorTo(pt4).Length
                'МСК
                Dim pt1 As Point2d = New Point2d(ArrayMSK(0, i), ArrayMSK(1, i))
                Dim pt2 As Point2d = New Point2d(ArrayMSK(0, j), ArrayMSK(1, j))
                Dim lenghtMSK As Double = pt1.GetVectorTo(pt2).Length
                'находим масштабный фактор
                Dim mXY As Double = lenghtMSK / lenghtSK
                ScaleFactor = ScaleFactor + mXY
                'добавляем строку в таблицу и записываем результат
                DataGridView3.Rows.Insert(CountScaleFactor)
                DataGridView3.Rows(CountScaleFactor).Cells(0).Value = True
                DataGridView3.Rows(CountScaleFactor).Cells(1).Value = ArrayMSK(2, i) & " - " & ArraySK(2, j)
                If CheckBox1.Checked = True Then
                    DataGridView3.Rows(CountScaleFactor).Cells(2).Value = 1
                Else
                    DataGridView3.Rows(CountScaleFactor).Cells(2).Value = Math.Round(mXY, 7)
                End If
                CountScaleFactor += 1
            Next j
        Next i
        'записываем масштабный фактор (усредненный)
        ScaleFactor = Math.Round(ScaleFactor / CountScaleFactor, 6)
        DataGridView3.Rows(CountScaleFactor).Cells(2).Value = Math.Round(ScaleFactor, 7) 'записываем в последнюю строку среднее значение
        DataGridView3.Rows(CountScaleFactor).Cells(1).Value = "Ср.знач" 'записываем в последнюю строку среднее значение
        '4.находим угол доворота
        Dim CountRotationAngle As Double = 0
        For i As Integer = 0 To ArraySK.GetUpperBound(1) - 1
            For j As Integer = i + 1 To ArraySK.GetUpperBound(1)
                'СК
                Dim pt3 As Point2d = New Point2d(ArraySK(0, i), ArraySK(1, i))
                Dim pt4 As Point2d = New Point2d(ArraySK(0, j), ArraySK(1, j))
                Dim RotationSK As Double = pt3.GetVectorTo(pt4).Angle
                'МСК
                Dim pt1 As Point2d = New Point2d(ArrayMSK(0, i), ArrayMSK(1, i))
                Dim pt2 As Point2d = New Point2d(ArrayMSK(0, j), ArrayMSK(1, j))
                Dim RotationMSK As Double = pt1.GetVectorTo(pt2).Angle
                'находим разность углов
                Dim AngleXY As Double = Math.Round(RotationSK - RotationMSK, 8)
                'If AngleXY < 0 Then AngleXY = AngleXY + Math.PI * 2
                'записываем результат в таблицу
                DataGridView3.Rows(CountRotationAngle).Cells(3).Value = FuncFormatZN.FuncFormatDEGToDMMSS((AngleXY * 180) / Math.PI)
                RotationAngle = RotationAngle + AngleXY
                CountRotationAngle += 1
            Next j
        Next i
        'записываем угол доворота (усредненный)
        RotationAngle = Math.Round(RotationAngle / CountRotationAngle, 6)
        DataGridView3.Rows(CountRotationAngle).Cells(3).Value = FuncFormatZN.FuncFormatDEGToDMMSS((RotationAngle * 180) / Math.PI)
        '5.находим коэффициенты a и b
        Dim A1 As Double = Math.Round(ScaleFactor * Math.Cos(RotationAngle), 7)
        Dim A2 As Double = -Math.Round(ScaleFactor * Math.Sin(RotationAngle), 7)
        Dim B1 As Double = Math.Round(ScaleFactor * Math.Sin(RotationAngle), 7)
        Dim B2 As Double = Math.Round(ScaleFactor * Math.Cos(RotationAngle), 7)
        '6. Считаем сдвиг осей координат
        Dim CountDXDY As Integer = 0
        For i As Integer = 0 To ArraySK.GetUpperBound(1) - 1
            For j As Integer = i + 1 To ArraySK.GetUpperBound(1)
                Dim C1 As Double = Math.Round((ArraySK(1, i) * A1 + ArraySK(0, i) * A2) - ArrayMSK(1, i), 4) 'смещение по x
                Dim C2 As Double = Math.Round((ArraySK(1, i) * B1 + ArraySK(0, i) * B2) - ArrayMSK(0, i), 4) 'смещение по y
                DataGridView3.Rows(CountDXDY).Cells(4).Value = C1
                DataGridView3.Rows(CountDXDY).Cells(5).Value = C2
                DX = DX + C1
                DY = DY + C2
                CountDXDY += 1
            Next j
        Next i
        'записываем смещение координат-линейный сдвиг (усредненный)
        DX = -1 * Math.Round(DX / CountDXDY, 4)
        DY = -1 * Math.Round(DY / CountDXDY, 4)
        DataGridView3.Rows(CountDXDY).Cells(4).Value = DX
        DataGridView3.Rows(CountDXDY).Cells(5).Value = DY
        '7.записываем результат
        DataGridView5.Rows.Insert(0, 6)
        DataGridView5.Rows(0).HeaderCell.Value = "A1"
        DataGridView5.Rows(0).Cells(0).Value = A1

        DataGridView5.Rows(1).HeaderCell.Value = "A2"
        DataGridView5.Rows(1).Cells(0).Value = B1

        DataGridView5.Rows(2).HeaderCell.Value = "B1"
        DataGridView5.Rows(2).Cells(0).Value = A2

        DataGridView5.Rows(3).HeaderCell.Value = "B2"
        DataGridView5.Rows(3).Cells(0).Value = B2

        DataGridView5.Rows(4).HeaderCell.Value = "A0"
        DataGridView5.Rows(4).Cells(0).Value = DY

        DataGridView5.Rows(5).HeaderCell.Value = "B0"
        DataGridView5.Rows(5).Cells(0).Value = DX
        '8. Делаем проверку вычисленных значений
        For i As Integer = 0 To ArraySK.GetUpperBound(1)
            Dim C1 As Double = Math.Round((ArraySK(1, i) * A1 + ArraySK(0, i) * A2) + DX - ArrayMSK(1, i), 4) 'смещение по x
            Dim C2 As Double = Math.Round((ArraySK(1, i) * B1 + ArraySK(0, i) * B2) + DY - ArrayMSK(0, i), 4) 'смещение по y
            DataGridView4.Rows.Insert(i)
            DataGridView4.Rows(i).Cells(0).Value = ArraySK(2, i)
            DataGridView4.Rows(i).Cells(1).Value = C1
            DataGridView4.Rows(i).Cells(2).Value = C2
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class