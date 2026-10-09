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

Public Class LoadsImageLinkedGPS
    '==============================================================================================
    'загружает гипперссылки по данным изображений с привязкой к GPS
    <CommandMethod("LoadsImageLinkedGPS")>
    Public Sub LoadsImageLinkedGPS()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
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
        Dim FormImageLincGPS As ImageLinkedGPS = New ImageLinkedGPS
        FormImageLincGPS.ShowDialog()
        If FormImageLincGPS.boolShowDialog = False Then
            Exit Sub
        End If
        Dim PutchFileTemplate As String = FormImageLincGPS.ComboBox2.Tag & FormImageLincGPS.ComboBox2.Text
        If File.Exists(PutchFileTemplate) = False Then
            MsgBox("Шаблон не найден. Макро прервано!!!")
            Exit Sub
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'загружаем образец блока
        Dim ArrayPropBlk As String(,) = Nothing
        Dim NameBlk As String = "Гиперссылка на геопривязанное изображение"
        Dim NameLayer As String = "0"
        Dim IndColor As Integer = 256
        Dim ScaleBlk As Double = 1
        Dim boolPropBlkUPK As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchFileTemplate, NameBlk, ArrayPropBlk, "Блок", "Иные (вспомогательные) объекты")
        If IsArray(ArrayPropBlk) = True Then
            For i As Integer = 0 To ArrayPropBlk.GetUpperBound(1)
                Dim NameField As String = ArrayPropBlk(0, i)
                If NameField.Trim Like "Name" Then
                    NameBlk = ArrayPropBlk(1, i)
                    If IsNothing(NameBlk) = False Then
                        If NameBlk.Trim.Length > 0 Then
                            Dim boolInsRefBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(PutchFileTemplate, NameBlk)
                            Exit For
                        Else
                            MsgBox("Определение блока в чертеж не загружено. Макро прервано!!!")
                            Exit Sub
                        End If
                    Else
                        MsgBox("Определение блока в чертеж не загружено. Макро прервано!!!")
                        Exit Sub
                    End If
                ElseIf NameField.Trim Like "Color" Then
                    IndColor = Val(ArrayPropBlk(1, i))
                ElseIf NameField.Trim Like "Layer" Then
                    Dim TempNameLayer As String = ArrayPropBlk(1, i)
                    Dim boolInsRefLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchFileTemplate, TempNameLayer)
                    If boolInsRefLayer = True Then
                        NameLayer = TempNameLayer
                    End If
                ElseIf NameField.Trim Like "Scale" Then
                    ScaleBlk = Val(ArrayPropBlk(1, i))
                End If
            Next i
        Else
            MsgBox("Определение блока в чертеж не загружено. Макро прервано!!!")
            Exit Sub
        End If
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '1.считывае координатную зону чертежа
            Dim CoordSystemCode As String = FormImageLincGPS.ComboBox1.Text
            Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
            Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
            Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
            Dim datumDict As MgCoordinateSystemDatumDictionary = csCatalog.GetDatumDictionary()
            Dim ellipsoidDict As MgCoordinateSystemEllipsoidDictionary = csCatalog.GetEllipsoidDictionary
            Dim LL84Cs As MgCoordinateSystem = csDict.GetCoordinateSystem("LL84")
            Dim DrawingSysCoord As MgCoordinateSystem = csDict.GetCoordinateSystem(CoordSystemCode)
            '2.выбираем изображения
            Dim fFile1 As String() = FormImageLincGPS.TextBox1.Tag
            '3. Выбираем результирующий каталог
            Dim FolderName As String = FormImageLincGPS.TextBox2.Text
            If Directory.Exists(FolderName) = False Then
                FolderName = ""
            End If
            '4. тело
            Dim LatWGS As Double = Nothing 'широта изображения в десятичных градусах
            Dim LonWGS As Double = Nothing 'долгота изображения в десятичных градусах
            Dim HGPS As Double = Nothing 'высота м метрах
            Dim FullNameFileLink As String = Nothing 'полный путь к файлу для гиперссылки
            'обсалютный или относительный путь
            Dim NameErrorFiles As String() = Nothing
            Dim ErrorCount As Integer = 0
            If IsArray(fFile1) = True Then
                For i As Integer = 0 To fFile1.Length - 1
                    Try
                        Dim image1 As System.Drawing.Bitmap = New System.Drawing.Bitmap(fFile1(i)) 'путь к файлу
                        'отделяем имя файла от пути
                        Dim pos As Integer = fFile1(i).LastIndexOf("\")
                        Dim nameFileImage As String = Mid(fFile1(i), pos + 1)
                        If FolderName <> "" Then
                            FileCopy(fFile1(i), FolderName & nameFileImage)
                            If FormImageLincGPS.RadioButton1.Checked = True Then
                                FullNameFileLink = FolderName & nameFileImage
                            Else
                                FullNameFileLink = nameFileImage
                            End If
                        Else
                            If FormImageLincGPS.RadioButton1.Checked = True Then
                                FullNameFileLink = fFile1(i)
                            Else
                                FullNameFileLink = nameFileImage
                            End If
                        End If
                        Dim propItems As System.Drawing.Imaging.PropertyItem() = image1.PropertyItems 'читаем все свойства
                        Dim propItem As System.Drawing.Imaging.PropertyItem = Nothing

                        Dim LatBytes As Byte() = Nothing 'широта
                        Dim LonBytes As Byte() = Nothing 'долгота
                        Dim HeightBytes As Byte() = Nothing 'высота
                        Dim boolLat As Boolean = False
                        Dim boolLon As Boolean = False
                        Dim boolHgt As Boolean = False
                        For Each propItem In propItems
                            Dim idProp As Integer = propItem.Id
                            If idProp = 2 Then 'считываем широту
                                If boolLat = False Then
                                    LatBytes = propItem.Value
                                    Dim grad As Integer = LatBytes(0) 'читаем значения целых градусов
                                    Dim min As Integer = LatBytes(8) 'читаем значения целых минут
                                    'получаем секунды
                                    Dim Bytesec As Byte() = {LatBytes(16), LatBytes(17), LatBytes(18), LatBytes(19), 0, 0, 0, 0}
                                    Dim sec As Decimal = BitConverter.ToUInt64(Bytesec, 0)
                                    'получаем размерность
                                    Dim Bytesec2 As Byte() = {LatBytes(20), LatBytes(21), LatBytes(22), LatBytes(23), 0, 0, 0, 0}
                                    Dim del As Decimal = BitConverter.ToUInt64(Bytesec2, 0)
                                    Dim sec1 As Double = sec / del
                                    LatWGS = grad + min / 60 + sec1 / 3600 '
                                    boolLat = True
                                End If
                            ElseIf idProp = 4 Then
                                If boolLon = False Then
                                    LonBytes = propItem.Value
                                    Dim grad As Integer = LonBytes(0)
                                    Dim min As Integer = LonBytes(8)
                                    Dim Bytesec As Byte() = {LonBytes(16), LonBytes(17), LonBytes(18), LonBytes(19), 0, 0, 0, 0}
                                    Dim sec As Decimal = BitConverter.ToUInt64(Bytesec, 0)
                                    Dim Bytesec2 As Byte() = {LonBytes(20), LonBytes(21), LonBytes(22), LonBytes(23), 0, 0, 0, 0}
                                    Dim del As Decimal = BitConverter.ToUInt64(Bytesec2, 0)
                                    Dim sec1 As Double = sec / del
                                    LonWGS = grad + min / 60 + sec1 / 3600
                                    boolLon = True
                                End If
                            ElseIf idProp = 6 Then
                                If boolHgt = False Then
                                    HeightBytes = propItem.Value
                                    Dim ByteH As Byte() = {HeightBytes(0), HeightBytes(1), 0, 0, 0, 0, 0, 0}
                                    HGPS = BitConverter.ToUInt64(ByteH, 0)
                                    boolHgt = True
                                End If
                            End If
                            If boolHgt = True And boolLat = True And boolLon = True Then Exit For
                        Next
                        ' преобразование из wgs84 в систему координат чертежа
                        Dim LL84WGS_to_DrawCoord As MgCoordinateSystemTransform = coordSysFactory.GetTransform(LL84Cs, DrawingSysCoord)
                        Dim LL84WGS_to_DrawCoordRez As MgCoordinate = LL84WGS_to_DrawCoord.Transform(LonWGS, LatWGS)
                        Dim xrez As Double = LL84WGS_to_DrawCoordRez.X
                        Dim yrez As Double = LL84WGS_to_DrawCoordRez.Y
                        'вставляем блок по этим координатам
                        '1.определяем точку вставки
                        Dim PositionBLK As Point3d = New Point3d(xrez, yrez, 0)
                        Dim acBLK As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(PositionBLK, ArrayPropBlk, 0)
                        If IsNothing(acBLK) = False Then
                            If acBLK.Id.IsNull = False Then
                                'находим id по Handle обукта
                                If acBLK.IsWriteEnabled = False Then
                                    acBLK.UpgradeOpen()
                                End If
                                Dim Hypln As HyperLink = New HyperLink
                                Hypln.Name = FullNameFileLink
                                acBLK.Hyperlinks.Add(Hypln)
                                image1.Dispose()
                                image1 = Nothing
                            End If
                        End If
                    Catch ex As ArgumentException
                        ReDim Preserve NameErrorFiles(ErrorCount)
                        NameErrorFiles(ErrorCount) = FullNameFileLink
                        ErrorCount += 1
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        ReDim Preserve NameErrorFiles(ErrorCount)
                        NameErrorFiles(ErrorCount) = FullNameFileLink
                        ErrorCount += 1
                    End Try
line1:
                Next i
            End If
            acTrans.Commit()
        End Using
LineExit:
    End Sub
End Class

