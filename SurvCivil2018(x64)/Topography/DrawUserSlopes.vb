Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Public Class DrawUserSlopes
    Inherits System.Windows.Forms.Form
    Public stbool As Boolean = False
    Public PutchImageObject As String = ""

    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTables As String = "d:\AutodeskCivilTools\FileResources\Sample\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTables = parentPatch2 & "\FileResources\Sample\"
        End If
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            ComboBox1.Tag = PutchTables
            Try
                For Each fileDWG As String In Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        ReDim Preserve ArrayTemplates(CountArrayTemplates)
                        ArrayTemplates(CountArrayTemplates) = Path.GetFileName(fileDWG)
                        CountArrayTemplates += 1
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
        If IsArray(ArrayTemplates) = True Then
            ComboBox1.DataSource = ArrayTemplates
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'загружаем изображение
        Dim PutchImageObject As String = "d:\AutodeskCivilTools\FileResources\ImageObject\"
        If IO.Directory.Exists(parentPatch2 & "\FileResources\ImageObject\") Then
            PutchImageObject = parentPatch2 & "\FileResources\ImageObject\"
            Try
                Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImageObject & "Откос неукрепленный.png")
                PictureBox1.Image = CType(MyImage, System.Drawing.Image)
                PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                ComboBox1.Text = "ИИ-ТОПО-Откос неукрепленный"
            Catch ex As ArgumentException
            End Try
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'Масштаб чертежа
        Dim AnnoScale As AnnotationScale = acCurDb.Cannoscale
        Dim MDraw As Double = (AnnoScale.PaperUnits / AnnoScale.Scale) * 1000
        NumericUpDown1.Value = MDraw
    End Sub
    'нажата кнопка НАРИСОВАТЬ ОТКОС
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim randomObject As New Random() 'формируем индивидуальный id для откоса
        Dim randomNumber As Integer = randomObject.Next()
        stbool = True
        Me.Hide()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'тип откоса
        Dim TipOtk As String = ""
        If RadioButton1.Checked = True Then
            TipOtk = "Откос неукрепленный"
        ElseIf RadioButton2.Checked = True Then
            TipOtk = "Откос укрепленный"
        ElseIf RadioButton3.Checked = True Then
            TipOtk = "Обрыв земляной"
        ElseIf RadioButton7.Checked = True Then
            TipOtk = "Стенка подпорная каменная"
        ElseIf RadioButton8.Checked = True Then
            TipOtk = "Стенка подпорная деревянная"
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acLyrTbl As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForWrite)
            Dim NameLayer As String = "0"
            Dim IndColor As Integer = 7
            Dim WLine As Double = 20
            Dim PatchTemplate As String = ComboBox1.Tag & ComboBox1.Text
            Dim ArrayPropLine As String(,) = Nothing
            If File.Exists(PatchTemplate) = True Then
                Dim boolPropLine As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PatchTemplate, TipOtk, ArrayPropLine, "Полилиния", "Элементы ситуации и рельефа")
                If IsArray(ArrayPropLine) = True Then
                    For i As Integer = 0 To ArrayPropLine.GetUpperBound(1)
                        Dim NameField As String = ArrayPropLine(0, i)
                        If NameField.Trim Like "Наименование слоя" Then
                            Dim TempNameLayer As String = ArrayPropLine(1, i)
                            If acLyrTbl.Has(TempNameLayer) = False Then
                                Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PatchTemplate, TempNameLayer)
                                If acLyrTbl.Has(TempNameLayer) = True Then
                                    NameLayer = TempNameLayer
                                End If
                            Else
                                NameLayer = TempNameLayer
                            End If
                        ElseIf NameField.Trim Like "Цвет" Then
                            IndColor = Val(ArrayPropLine(1, i))
                            If IndColor < 0 Then IndColor = 0
                            If IndColor > 256 Then IndColor = 256
                        ElseIf NameField.Trim Like "Вес линии" Then
                            WLine = Val(ArrayPropLine(1, i))
                        End If
                    Next i
                Else
                    MsgBox("Таблица не найдена. Макро прервано!!!")
                    Exit Sub
                End If

            Else
                MsgBox("Шаблон не найден. Макро прервано!!!")
                Exit Sub
            End If

            'масштаб топоплана
            Dim M As Double = NumericUpDown1.Value
            'длина короткого штриха
            Dim LenghtSh As Double = NumericUpDown2.Value * (M / 1000)
            'расстояние между штрихами длинными
            Dim DistSh As Double = NumericUpDown3.Value * (M / 1000)
            'короткий штрих равен половине длинного
            Dim boolshSr As Boolean = False
            If CheckBox3.Checked = True Then boolshSr = True
            'запускаем процедуру отрисовки откоса
            Dim otkosDraw As DrawSlopes = New DrawSlopes
            If RadioButton4.Checked = True Then
                otkosDraw.AutoDrawSlope(LenghtSh, TipOtk, NameLayer, IndColor, DistSh, randomNumber, M, boolshSr, WLine) 'перепендикуляр от бровки
            ElseIf RadioButton9.Checked = True Then
                Dim otkosDel As DeleteSlopes = New DeleteSlopes
                otkosDel.DeleteSlopes()
            Else
                otkosDraw.ManuallyDrawSlope(LenghtSh, TipOtk, NameLayer, IndColor, DistSh) 'построение в ручную
            End If
            acTrans.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        stbool = 0
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'неукрепленный откос
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If Directory.Exists(PutchImageObject) = True Then
            Try
                Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImageObject & "Откос неукрепленный.png")
                PictureBox1.Image = CType(MyImage, System.Drawing.Image)
                PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Catch ex As ArgumentException
            End Try
            NumericUpDown3.Value = 1 'расстояние между штрихами
            NumericUpDown2.Value = 2 'длина короткого штриха
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'укрепленный откос
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If Directory.Exists(PutchImageObject) = True Then
            Try
                Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImageObject & "Откос укрепленный.png")
                PictureBox1.Image = CType(MyImage, System.Drawing.Image)
                PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Catch ex As ArgumentException
            End Try
            NumericUpDown3.Value = 1 'расстояние между штрихами
            NumericUpDown2.Value = 2 'длина короткого штриха
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If Directory.Exists(PutchImageObject) = True Then
            Try
                Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImageObject & "Обрыв земляной.png")
                PictureBox1.Image = CType(MyImage, System.Drawing.Image)
                PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Catch ex As ArgumentException
            End Try
            NumericUpDown3.Value = 1 'расстояние между штрихами
            NumericUpDown2.Value = 0 'длина короткого штриха
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If Directory.Exists(PutchImageObject) = True Then
            Try
                Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImageObject & "Подпорная стенка каменная или бетонная.png")
                PictureBox1.Image = CType(MyImage, System.Drawing.Image)
                PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Catch ex As ArgumentException
            End Try
            NumericUpDown3.Value = 4 'расстояние между штрихами
            NumericUpDown2.Value = 0 'длина короткого штриха
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If Directory.Exists(PutchImageObject) = True Then
            Try
                Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImageObject & "Подпорная стенка деревянная.png")
                PictureBox1.Image = CType(MyImage, System.Drawing.Image)
                PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
            Catch ex As ArgumentException
            End Try
            NumericUpDown3.Value = 4 'расстояние между штрихами
            NumericUpDown2.Value = 0 'длина короткого штриха
        End If

    End Sub

    Private Sub DrawUserSlopes_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class