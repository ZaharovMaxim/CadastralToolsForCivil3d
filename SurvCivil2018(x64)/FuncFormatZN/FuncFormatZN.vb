Imports System.Globalization
Imports Autodesk.AutoCAD.Geometry
Imports Microsoft.Office.Interop

Public Class FuncFormatZN
    '==================================================================================================================================
    'функция форматирует длины линий под определенный формат
    Public Shared Function FuncFormatCoordinate(ByVal coordinates As Double, Optional ByVal RountZn As Integer = 2) As String
        FuncFormatCoordinate = coordinates
        Dim Rn As Integer = RountZn + 3
        For i As Integer = Rn To RountZn Step -1
            coordinates = Math.Round(coordinates, i)
        Next
        Try
            If RountZn = 0 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0}", Math.Round(coordinates, RountZn))
            ElseIf RountZn = 1 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Round(coordinates, RountZn))
            ElseIf RountZn = 2 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Round(coordinates, RountZn))
            ElseIf RountZn = 3 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Round(coordinates, RountZn))
            ElseIf RountZn = 4 Then
                FuncFormatCoordinate = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Round(coordinates, RountZn))
            Else
                FuncFormatCoordinate = Math.Round(coordinates, RountZn)
            End If
        Catch
            Exit Function
        End Try
    End Function
    '==========================================================================
    'функция преобразут ввденные пользователем значения градусов в десятичном формате в градусы0 мин'сек"
    Public Shared Function FuncConvertDegToDDMMSS(ByVal DegGrad As Double, Optional ByVal units As Integer = 1, Optional ByVal roudZn As Integer = 0) As String
        FuncConvertDegToDDMMSS = "0"
        If units = 1 Then
            roudZn = roudZn + 4
        End If
        Dim StrAngle As String = Autodesk.AutoCAD.Runtime.Converter.AngleToString(DegGrad, units, roudZn)
        If units = 1 Then
            StrAngle = StrAngle.Replace("d", Chr(176))
        End If
        FuncConvertDegToDDMMSS = StrAngle
    End Function
    '==================================================================================================================================
    'функция форматирует десятичное число в ПК+(с учетом количества знаков после запятой
    Public Shared Function FuncFormatPK(ByVal DecimalPK As Double, Optional ByVal lenStr As Integer = 2) As String
        FuncFormatPK = "0+00.00"
        Try
            'отделяем число полных пикетов
            Dim IntDist As Integer = Int(DecimalPK / 100)
            'отделяем плюсовую часть
            Dim PKPlus As Double = DecimalPK - IntDist * 100
            If PKPlus = 100 Then
                PKPlus = 0
                IntDist = IntDist + 1
            End If
            'форматируем число (целое число пикетов)
            Dim StrPKPlus As String = String.Format("{0:00.00}", PKPlus)
            FuncFormatPK = IntDist & "+" & StrPKPlus
        Catch
            Exit Function
        End Try
    End Function
    '==================================================================================================================================
    'функция форматирует десятичное число в километр+
    Public Shared Function FuncFormatKM(ByVal DecimalKM As Double, Optional ByVal lenStr As Integer = 2) As String
        FuncFormatKM = "0+000.00"
        Try
            'отделяем число полных километров
            Dim IntDist As Integer = Int(DecimalKM / 1000)
            'отделяем плюсовую часть
            Dim KMPlus As Double = Math.Round(DecimalKM - IntDist * 1000, 2)
            If KMPlus = 1000 Then
                KMPlus = 0
                IntDist = IntDist + 1
            End If
            'преобразуем число в строку
            Dim StrKMPlus As String = String.Format("{0:000.00}", KMPlus)
            FuncFormatKM = IntDist & "+" & StrKMPlus
        Catch
            Exit Function
        End Try
    End Function
    '==========================================================================
    'функция преобразут ввденные пользователем значения градусов в десятичном формате в градусы мин сек (простое форматирование)
    Public Shared Function FuncFormatDEGToDMMSS(ByVal Gragus As Double) As String
        FuncFormatDEGToDMMSS = "0.0000"
        Try
            Dim znk As String = ""
            If Gragus < 0 Then
                znk = "-"
                Gragus = Math.Abs(Gragus)
            End If
            Dim Gr As Double = Int(Gragus)
            Dim dr As Double = Math.Round(Gragus - Gr, 4)
            Dim dr_str As String = dr.ToString
            Dim lenStr As Integer = dr_str.Trim.Length
            If lenStr = 1 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & ".0000"
            ElseIf lenStr = 3 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & "000"
            ElseIf lenStr = 4 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & "00"
            ElseIf lenStr = 5 Then
                FuncFormatDEGToDMMSS = Gragus.ToString & "0"
            Else
                FuncFormatDEGToDMMSS = Gragus.ToString
            End If
        Catch
            FuncFormatDEGToDMMSS = "0.0000"
            Exit Function
        End Try
    End Function

    Public Shared Function FuncConvertStrNumberPoint(ByVal xStr As String) As String
        Dim ii, i1, i2, i3, i4, aa, aaa, oStr
        aa = Split(xStr, ",") 'разбиваем строку на массив подстрок по разделителю - запятая
        aaa = Split(aa(0), "-") ' разбиваем начальный элемент этого массива подстрок по разделителю - тире. 
        i1 = CInt(aaa(0)) : i2 = CInt(aaa(UBound(aaa))) 'первое число -> i1, последнее (если оно одно, то первое) -> i2
        oStr = CStr(i1) ' запишем в результирующую строку пока самое первое число
        For ii = 1 To UBound(aa) Step 1 ' пробежимся по остальным элементам массива подстрок
            aaa = Split(aa(ii), "-") ' каждый разбиваем на массив полуподстрок - получаем 1 или 2 элемента
            i3 = CInt(aaa(0)) : i4 = CInt(aaa(UBound(aaa))) 'первое число -> i3, последнее (если оно одно - оно же и первое) -> i4
            If i3 = i2 + 1 Then ' Если это смежные диапазоны чисел 
                i2 = i4         ' - сольём их в один
            Else              ' в противном случае дозаписываем хвост предыдущего диапазона
                If i2 = i1 + 1 Then oStr = oStr & "," & CStr(i2) 'если диапазон всего из 2 чисел
                If i2 > i1 + 1 Then oStr = oStr & "-" & CStr(i2) 'если диапазон поболее
                i1 = i3 : i2 = i4 'текущий диапазон становится прежним
                oStr = oStr & "," & CStr(i1) 'Дозаписываем начало этого диапазона в результирующую строку 
            End If
        Next
        If i2 = i1 + 1 Then oStr = oStr & "," & CStr(i2) 'дозаписываем хвост последнего диапазаноа
        If i2 > i1 + 1 Then oStr = oStr & "-" & CStr(i2)
        FuncConvertStrNumberPoint = oStr ' результат возвращаем
    End Function

    '================================================================================================================================
    'функция декодирует символы Unicode
    Public Shared Function FuncUnicodDecode(ByVal str As String) As String
        Dim Tempstr As String() = str.Split("\")
        Dim FinalStr As String = ""
        For i As Integer = 0 To Tempstr.Length - 1
            Dim tstr As String = Tempstr(i)
            If Left(tstr, 2) Like "u0" Then
                Dim tstr2 As String = Mid(tstr, 2, 4)
                Select Case tstr2.ToUpper
                    Case "0410"
                        FinalStr = FinalStr & "А"
                    Case "0411"
                        FinalStr = FinalStr & "Б"
                    Case "0412"
                        FinalStr = FinalStr & "В"
                    Case "0413"
                        FinalStr = FinalStr & "Г"
                    Case "0414"
                        FinalStr = FinalStr & "Д"
                    Case "0415"
                        FinalStr = FinalStr & "Е"
                    Case "0416"
                        FinalStr = FinalStr & "Ж"
                    Case "0417"
                        FinalStr = FinalStr & "З"
                    Case "0418"
                        FinalStr = FinalStr & "И"
                    Case "0419"
                        FinalStr = FinalStr & "Й"
                    Case "041A"
                        FinalStr = FinalStr & "К"
                    Case "041B"
                        FinalStr = FinalStr & "Л"
                    Case "041C"
                        FinalStr = FinalStr & "М"
                    Case "041D"
                        FinalStr = FinalStr & "Н"
                    Case "041E"
                        FinalStr = FinalStr & "О"
                    Case "041F"
                        FinalStr = FinalStr & "П"
                    Case "0420"
                        FinalStr = FinalStr & "Р"
                    Case "0421"
                        FinalStr = FinalStr & "С"
                    Case "0422"
                        FinalStr = FinalStr & "Т"
                    Case "0423"
                        FinalStr = FinalStr & "У"
                    Case "0424"
                        FinalStr = FinalStr & "Ф"
                    Case "0425"
                        FinalStr = FinalStr & "Х"
                    Case "0426"
                        FinalStr = FinalStr & "Ц"
                    Case "0427"
                        FinalStr = FinalStr & "Ч"
                    Case "0428"
                        FinalStr = FinalStr & "Ш"
                    Case "0429"
                        FinalStr = FinalStr & "Щ"
                    Case "042A"
                        FinalStr = FinalStr & "Ъ"
                    Case "042B"
                        FinalStr = FinalStr & "Ы"
                    Case "042C"
                        FinalStr = FinalStr & "Ь"
                    Case "042D"
                        FinalStr = FinalStr & "Э"
                    Case "042E"
                        FinalStr = FinalStr & "Ю"
                    Case "042F"
                        FinalStr = FinalStr & "Я"

                    Case "0430"
                        FinalStr = FinalStr & "а"
                    Case "0431"
                        FinalStr = FinalStr & "б"
                    Case "0432"
                        FinalStr = FinalStr & "в"
                    Case "0433"
                        FinalStr = FinalStr & "г"
                    Case "0434"
                        FinalStr = FinalStr & "д"
                    Case "0435"
                        FinalStr = FinalStr & "е"
                    Case "0436"
                        FinalStr = FinalStr & "ж"
                    Case "0437"
                        FinalStr = FinalStr & "з"
                    Case "0438"
                        FinalStr = FinalStr & "и"
                    Case "0439"
                        FinalStr = FinalStr & "й"
                    Case "043A"
                        FinalStr = FinalStr & "к"
                    Case "043B"
                        FinalStr = FinalStr & "л"
                    Case "043C"
                        FinalStr = FinalStr & "м"
                    Case "043D"
                        FinalStr = FinalStr & "н"
                    Case "043E"
                        FinalStr = FinalStr & "о"
                    Case "043F"
                        FinalStr = FinalStr & "п"
                    Case "0440"
                        FinalStr = FinalStr & "р"
                    Case "0441"
                        FinalStr = FinalStr & "с"
                    Case "0442"
                        FinalStr = FinalStr & "т"
                    Case "0443"
                        FinalStr = FinalStr & "у"
                    Case "0444"
                        FinalStr = FinalStr & "ф"
                    Case "0445"
                        FinalStr = FinalStr & "х"
                    Case "0446"
                        FinalStr = FinalStr & "ц"
                    Case "0447"
                        FinalStr = FinalStr & "ч"
                    Case "0448"
                        FinalStr = FinalStr & "ш"
                    Case "0449"
                        FinalStr = FinalStr & "щ"
                    Case "044A"
                        FinalStr = FinalStr & "ъ"
                    Case "044B"
                        FinalStr = FinalStr & "ы"
                    Case "044C"
                        FinalStr = FinalStr & "ь"
                    Case "044D"
                        FinalStr = FinalStr & "э"
                    Case "044E"
                        FinalStr = FinalStr & "ю"
                    Case "044F"
                        FinalStr = FinalStr & "я"

                    Case "0022"
                        FinalStr = FinalStr & Chr(34)
                    Case "0028"
                        FinalStr = FinalStr & "("
                    Case "0029"
                        FinalStr = FinalStr & ")"
                    Case "002C"
                        FinalStr = FinalStr & ","
                    Case "002E"
                        FinalStr = FinalStr & "-"
                    Case "002F"
                        FinalStr = FinalStr & "/"
                    Case "0030"
                        FinalStr = FinalStr & "0"
                    Case "0031"
                        FinalStr = FinalStr & "1"
                    Case "0032"
                        FinalStr = FinalStr & "2"
                    Case "0033"
                        FinalStr = FinalStr & "3"
                    Case "0034"
                        FinalStr = FinalStr & "4"
                    Case "0035"
                        FinalStr = FinalStr & "5"
                    Case "0036"
                        FinalStr = FinalStr & "6"
                    Case "0037"
                        FinalStr = FinalStr & "7"
                    Case "0038"
                        FinalStr = FinalStr & "8"
                    Case "0039"
                        FinalStr = FinalStr & "9"
                    Case "003A"
                        FinalStr = FinalStr & ":"
                    Case "003B"
                        FinalStr = FinalStr & ";"
                    Case Else
                        FinalStr = FinalStr & "#"
                End Select
                If tstr.Length > 4 Then
                    FinalStr = FinalStr & Mid(tstr, 6)
                End If
            ElseIf Left(tstr, 1) Like "{" Then
                Continue For
            ElseIf Left(tstr, 1) Like "C" Then
                Dim pos As Integer = tstr.IndexOf(";")
                If pos > 0 Then
                    Dim tempArraystr As String() = tstr.Split(";")
                    If IsArray(tempArraystr) = True Then
                        If tempArraystr.Length > 0 Then
                            FinalStr = FinalStr & tempArraystr(tempArraystr.Length - 1)
                        End If
                    End If
                End If
            Else
                FinalStr = FinalStr & tstr
            End If
        Next
        FinalStr = FinalStr.Trim
        If Right(FinalStr, 1) Like "}" Then
            FinalStr = Mid(FinalStr, 1, FinalStr.Length - 1)
        End If
        FuncUnicodDecode = FinalStr
    End Function

    '==========================================================================
    'функция преобразует направление вектора в румб и форматирует его
    Public Shared Function FuncAngleDirectFormat(ByVal pt1 As Point2d, ByVal pt2 As Point2d, Optional ByVal ddmmss As Integer = 1, Optional ByVal roudZn As Integer = 0) As String
        Dim Direction As Double = 0
        Try
            Dim vk As Vector2d = pt1.GetVectorTo(pt2)
            Dim angleRad As Double = vk.Angle
            Dim AngleDeg As Double = (angleRad * 180) / (Math.PI)
            Select Case AngleDeg
                Case 0
                    Direction = 90
                Case 90
                    Direction = 0
                Case 180
                    Direction = 270
                Case 270
                    Direction = 180
                Case 0 To 90
                    Direction = 90 - AngleDeg

                Case 90 To 180
                    Direction = 450 - AngleDeg

                Case 180 To 270
                    Direction = 450 - AngleDeg

                Case 270 To 360
                    Direction = 450 - AngleDeg

            End Select
            If ddmmss = 0 Then
                FuncAngleDirectFormat = Math.Round(Direction, roudZn) & Chr(176)
            ElseIf ddmmss = 1 Then
                Dim Grad As Double = Int(Direction)
                Dim Min As Double = Math.Round((Direction - Grad) * 60, roudZn)
                Dim Min2 As String = "00"
                If Min < 10 Then
                    Min2 = "0" & Min
                Else
                    Min2 = Min
                End If
                FuncAngleDirectFormat = Grad & Chr(176) & Min2 & "'"
            Else
                Dim Grad As Double = Int(Direction)
                Dim Min As Double = Int((Direction - Grad) * 60)
                Dim Min2 As String = "00"
                If Min < 10 Then
                    Min2 = "0" & Min
                Else
                    Min2 = Min
                End If
                Dim Sec As Double = Math.Round(((Direction - Grad) * 60 - Min) * 60, roudZn)
                Dim Sec2 As String = "00"
                If Sec < 10 Then
                    Sec2 = "0" & Sec
                Else
                    Sec2 = Sec
                End If
                FuncAngleDirectFormat = Grad & Chr(176) & Min2 & "'" & Sec2 & """"
            End If
        Catch
            FuncAngleDirectFormat = "0.0000"
            Exit Function
        End Try
    End Function
    '==========================================================================
    'функция преобразует направление вектора в румб и форматирует его
    Public Shared Function FuncAngleRumbFormat(ByVal pt1 As Point2d, ByVal pt2 As Point2d, Optional ByVal ddmmss As Integer = 1, Optional ByVal roudZn As Integer = 0) As String
        Dim Rumb As Double = 0
        Dim Pref As String = ""
        Try
            Dim vk As Vector2d = pt1.GetVectorTo(pt2)
            Dim angleRad As Double = vk.Angle
            Dim AngleDeg As Double = (angleRad * 180) / (Math.PI)
            Select Case AngleDeg
                Case 0
                    FuncAngleRumbFormat = "В 0" & Chr(176)
                    Pref = "В"
                Case 90
                    FuncAngleRumbFormat = "С 0" & Chr(176)
                    Pref = "С"
                Case 180
                    FuncAngleRumbFormat = "З 0" & Chr(176)
                    Pref = "З"
                Case 270
                    FuncAngleRumbFormat = "Ю 0" & Chr(176)
                    Pref = "Ю"
                Case 0 To 90
                    Rumb = 90 - AngleDeg
                    Pref = "СВ"
                Case 90 To 180
                    Rumb = AngleDeg - 90
                    Pref = "СЗ"
                Case 180 To 270
                    Rumb = 270 - AngleDeg
                    Pref = "ЮЗ"
                Case 270 To 360
                    Rumb = AngleDeg - 270
                    Pref = "ЮВ"
            End Select
            If ddmmss = 0 Then
                FuncAngleRumbFormat = Pref & " " & Math.Round(Rumb, roudZn) & Chr(176)
            ElseIf ddmmss = 1 Then
                Dim Grad As Double = Int(Rumb)
                Dim Min As Double = Math.Round((Rumb - Grad) * 60, roudZn)
                Dim Min2 As String = "00"
                If Min < 10 Then
                    Min2 = "0" & Min
                Else
                    Min2 = Min
                End If
                FuncAngleRumbFormat = Pref & " " & Grad & Chr(176) & Min2 & "'"
            Else
                Dim Grad As Double = Int(Rumb)
                Dim Min As Double = Int((Rumb - Grad) * 60)
                Dim Min2 As String = "00"
                If Min < 10 Then
                    Min2 = "0" & Min
                Else
                    Min2 = Min
                End If
                Dim Sec As Double = Math.Round(((Rumb - Grad) * 60 - Min) * 60, roudZn)
                Dim Sec2 As String = "00"
                If Sec < 10 Then
                    Sec2 = "0" & Sec
                Else
                    Sec2 = Sec
                End If
                FuncAngleRumbFormat = Pref & " " & Grad & Chr(176) & Min2 & "'" & Sec2 & """"
            End If
        Catch
            FuncAngleRumbFormat = "0.0000"
            Exit Function
        End Try
    End Function
    Public Shared Function FormatTableRow(ByVal table As Word.Table, ByVal rowIndex As Integer, Optional nameFont As String = "Times New Roman", Optional sizeFont As Integer = 10, Optional alignCenter As Boolean = False) As Boolean
        Try
            ' Проверяем, существует ли строка с таким индексом
            If table Is Nothing Then
                Return False
            End If
            If rowIndex < 1 OrElse rowIndex > table.Rows.Count Then
                Return False
            End If
            ' Получаем целевой диапазон (всю строку)
            Dim targetRow As Word.Row = table.Rows(rowIndex)
            Dim targetRange As Word.Range = targetRow.Range
            ' Применяем шрифт
            targetRange.Font.Name = nameFont
            ' Применяем размер шрифта (10 пунктов)
            If sizeFont > 0 Then
                targetRange.Font.Size = sizeFont
            End If
            'Применяем выравнивание по центру
            If alignCenter = True Then
                targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                ' Дополнительно: выравнивание по вертикали в каждой ячейке (опционально)
                For Each cell As Word.Cell In targetRow.Cells
                    cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                Next
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
End Class
