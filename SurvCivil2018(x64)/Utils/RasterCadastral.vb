Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System.Drawing

Public Class RasterCadastral
    <CommandMethod("IMGWJPW")>
    Public Sub InsertImageWithJpw()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Выбор файла изображения
        Dim ofd As New Autodesk.AutoCAD.Windows.OpenFileDialog("Выберите изображение", Nothing, "jpg;png;bmp;jpeg", "Image Files", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple)
        If ofd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return
        Dim arrayFiles As String() = ofd.GetFilenames()
        For Each imagePath As String In arrayFiles
            Dim jpwPath As String = Path.ChangeExtension(imagePath, ".jgw")
            If Not File.Exists(jpwPath) Then
                ed.WriteMessage(vbLf & "Ошибка: Файл привязки .jpw не найден в папке с изображением!")
                Return
            End If

            ' 1. Читаем мировые координаты из JPW файла
            Dim worldCoords As JpwData = ReadJpwFile(jpwPath)
            If worldCoords Is Nothing Then Return

            ' 2. Определяем размеры в пикселях через System.Drawing
            Dim bmp As Bitmap = Nothing
            Try
                bmp = New Bitmap(imagePath) ' Загружаем изображение, чтобы узнать его реальные размеры
                Dim pixelWidth As Double = bmp.Width
                Dim pixelHeight As Double = bmp.Height

                ' Вычисляем физические размеры в мировых единицах
                Dim widthInWorld As Double = pixelWidth * Math.Abs(worldCoords.PixelSizeX)
                Dim heightInWorld As Double = pixelHeight * Math.Abs(worldCoords.PixelSizeY)

                Using trans As Transaction = db.TransactionManager.StartTransaction()
                    ' Создаем определение изображения
                    Dim imgDef As RasterImageDef = New RasterImageDef()
                    imgDef.SourceFileName = imagePath
                    imgDef.Load() ' Загружаем данные в кэш AutoCAD

                    Dim imgDictId As ObjectId = RasterImageDef.GetImageDictionary(db)
                    If imgDictId.IsNull Then
                        imgDictId = RasterImageDef.CreateImageDictionary(db)
                    End If

                    Dim imgDict As DBDictionary = trans.GetObject(imgDictId, OpenMode.ForWrite)
                    Dim defName As String = Path.GetFileNameWithoutExtension(imagePath)
                    Dim imgDefId As ObjectId = imgDict.SetAt(defName, imgDef)
                    trans.AddNewlyCreatedDBObject(imgDef, True)

                    ' Создаем ссылку на изображение в чертеже
                    Dim imgRef As New RasterImage()
                    imgRef.ImageDefId = imgDefId
                    imgRef.ImageTransparency = True
                    imgRef.ShowImage = True

                    ' Устанавливаем геометрию (масштаб и поворот)
                    Dim angleRad As Double = worldCoords.Rotation * Math.PI / 180.0
                    ' Вектор оси X (ширина)
                    Dim uCorner As New Vector3d(widthInWorld * Math.Cos(angleRad), widthInWorld * Math.Sin(angleRad), 0)
                    ' Вектор оси Y (высота) — перпендикулярен оси X
                    Dim vOnPlane As New Vector3d(-heightInWorld * Math.Sin(angleRad), heightInWorld * Math.Cos(angleRad), 0)
                    ' Точка вставки (нижний левый угол)
                    Dim insertPt As New Point3d(worldCoords.X, worldCoords.Y - heightInWorld, 0)

                    imgRef.Orientation = New CoordinateSystem3d(insertPt, uCorner, vOnPlane)
                    imgRef.ShowImage = False
                    ' Добавляем в пространство модели
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    ms.AppendEntity(imgRef)
                    trans.AddNewlyCreatedDBObject(imgRef, True)

                    ' Связываем ссылку с определением
                    RasterImage.EnableReactors(True)
                    imgRef.AssociateRasterDef(imgDef)

                    trans.Commit()
                    ed.WriteMessage(vbLf & $"Изображение '{Path.GetFileName(imagePath)}' успешно вставлено!")
                End Using
            Catch ex As Exception
                ed.WriteMessage(vbLf & $"Ошибка: {ex.Message}")
            Finally
                If bmp IsNot Nothing Then bmp.Dispose()
            End Try
        Next
    End Sub

    ' Структура для данных из JPW
    Private Class JpwData
        Public Property PixelSizeX As Double
        Public Property PixelSizeY As Double
        Public Property X As Double
        Public Property Y As Double
        Public Property Rotation As Double
    End Class

    ' Функция чтения JPW файла
    Private Function ReadJpwFile(path As String) As JpwData
        Try
            Dim lines As String() = File.ReadAllLines(path)
            If lines.Length < 6 Then Return Nothing

            ' Парсинг чисел (учитываем точку или запятую как разделитель)
            'Dim psX As Double = Double.Parse(lines(0).Replace(".", ","))
            'Dim rot1 As Double = Double.Parse(lines(1).Replace(".", ",")) ' Обычно 0
            'Dim rot2 As Double = Double.Parse(lines(2).Replace(".", ",")) ' Обычно 0
            'Dim psY As Double = Double.Parse(lines(3).Replace(".", ","))
            'Dim xCoord As Double = Double.Parse(lines(4).Replace(".", ",").Replace(",", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator))
            'Dim yCoord As Double = Double.Parse(lines(5).Replace(".", ",").Replace(",", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator))
            Dim psX As Double = Val(lines(0))
            Dim rot1 As Double = Val(lines(1)) ' Обычно 0
            Dim rot2 As Double = Val(lines(2)) ' Обычно 0
            Dim psY As Double = Val(lines(3))
            Dim xCoord As Double = Val(lines(4))
            Dim yCoord As Double = Val(lines(5))
            ' Расчет угла поворота
            Dim rotationAngle As Double = 0
            If psX <> 0 Then
                rotationAngle = Math.Atan2(rot2, psX) * 180.0 / Math.PI
            End If

            Return New JpwData() With {
                    .PixelSizeX = psX,
                    .PixelSizeY = psY,
                    .X = xCoord,
                    .Y = yCoord,
                    .Rotation = rotationAngle
                }
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class
