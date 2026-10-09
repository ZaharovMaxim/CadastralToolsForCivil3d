Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput


Public Class HatchContour
    Inherits System.Windows.Forms.Form
    Public stbool As Boolean = False
    Public boolSelectBlk As Boolean = False
    Public idBlk As ObjectId = Nothing
    Public PutchImage As String = "d:\Civil3DToolsUtility\FileResources\ImageObject\"
    Public PutchBLK As String = "d:\Civil3DToolsUtility\FileResources\Sample\" 'пути для точечных условных знаков
    Public PutchMVBLK As String = "d:\Civil3DToolsUtility\FileResources\3d_Object\" 'пути для точечных условных знаков
    'начальный вызов формы
    Public Sub New()
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'находим папки для блоков, линий и подписей
        '1. путь к блокам
        If IO.Directory.Exists(parentPatch2) = True Then
            PutchImage = parentPatch2 & "\FileResources\ImageObject\"
            PictureBox1.Tag = PutchImage
        End If
        Try
            Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & "Штриховка Сенокосы.jpg")
            PictureBox1.Image = CType(MyImage, System.Drawing.Image)
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Catch ex As ArgumentException
        End Try
        '===========================================================================================
        'выбираем масштаб чертежа
        Dim AnnoScale As AnnotationScale = acCurDb.Cannoscale
        Dim MDraw As Double = (AnnoScale.PaperUnits / AnnoScale.Scale) * 1000
        NumericUpDown3.Value = MDraw
        'загружаем образцы блоков
        If IO.Directory.Exists(parentPatch2) = True Then
            PutchBLK = parentPatch2 & "\FileResources\Sample\"
            PutchMVBLK = parentPatch2 & "\FileResources\3d_Object\"
        End If
        'читаем шаблоны условных знаков
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        Try
            For Each fileDWG As String In Directory.GetFiles(PutchBLK)
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
        If IsArray(ArrayTemplates) = True Then
            ComboBox3.DataSource = ArrayTemplates
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем вкладку с блоками
        Dim ArrayDeskBlk As String() = Nothing
        Dim NameTable As String = "Растительность"
        Dim PutshTemplateFile As String = PutchBLK & ComboBox3.Text
        If File.Exists(PutshTemplateFile) = True Then
            'читаем все названия блоков с таблицы Растительность
            Dim boolDesk As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutshTemplateFile, NameTable, ArrayDeskBlk, "Блок")
            If IsArray(ArrayDeskBlk) = True Then
                ComboBox2.DataSource = ArrayDeskBlk
            End If
        End If
        Dim DeskBlk As String = ComboBox2.Text
        Dim ArrayNameBlk As String(,) = Nothing
        Dim boolName As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplateFile, DeskBlk, ArrayNameBlk, "Блок", NameTable)
        Dim NameImageZn As String = ""
        If IsArray(ArrayNameBlk) = True Then
            For i As Integer = 0 To ArrayNameBlk.GetUpperBound(1)
                If ArrayNameBlk(0, i) Like "Отображение" Then
                    NameImageZn = ArrayNameBlk(1, i)
                End If
            Next
        End If
        If IsNothing(NameImageZn) = False Then
            If NameImageZn.Trim.Length > 0 Then
                Try
                    Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & NameImageZn & ".png")
                    PictureBox2.Image = CType(MyImage, System.Drawing.Image)
                    PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                Catch ex As FileNotFoundException
                Catch ex As ArgumentException
                Catch ex As IOException
                End Try
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'растительность
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            If CheckBox1.Checked = False Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'заполняем вкладку с блоками
                Dim ArrayDeskBlk As String() = Nothing
                Dim NameTable As String = "Растительность"
                Dim PutshTemplateFile As String = PutchBLK & ComboBox3.Text
                If File.Exists(PutshTemplateFile) = True Then
                    'читаем все названия блоков с таблицы Растительность
                    Dim boolDesk As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutshTemplateFile, NameTable, ArrayDeskBlk, "Блок")
                    If IsArray(ArrayDeskBlk) = True Then
                        ComboBox2.DataSource = ArrayDeskBlk
                        boolSelectBlk = False
                    End If
                End If
                Dim DeskBlk As String = ComboBox2.Text
                Dim ArrayNameBlk As String(,) = Nothing
                Dim boolName As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplateFile, DeskBlk, ArrayNameBlk, "Блок", NameTable)
                Dim NameImageZn As String = ""
                If IsArray(ArrayNameBlk) = True Then
                    For i As Integer = 0 To ArrayNameBlk.GetUpperBound(1)
                        If ArrayNameBlk(0, i) Like "Отображение" Then
                            NameImageZn = ArrayNameBlk(1, i)
                        End If
                    Next
                End If
                Dim PutchImage As String = PictureBox1.Tag
                If IsNothing(NameImageZn) = False Then
                    If NameImageZn.Trim.Length > 0 Then
                        Try
                            Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & NameImageZn & ".png")
                            PictureBox2.Image = CType(MyImage, System.Drawing.Image)
                            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                        Catch ex As FileNotFoundException
                        Catch ex As ArgumentException
                        Catch ex As IOException
                        End Try
                    End If
                End If
            Else
                '3d блоки
                If IO.File.Exists(PutchMVBLK & "MVBlocks.dwg") = True Then
                    'читаем шаблоны условных знаков
                    Dim ArrayBlock3d As String() = Nothing
                    Dim ArrayNameTables As String() = {"Растительность"}
                    Dim boolfind3d As Boolean = AcadFunction.FuncReadNameObjectAllDestDWG(PutchMVBLK & "MVBlocks.dwg", ArrayNameTables, ArrayBlock3d, "MVBlock")
                    If IsArray(ArrayBlock3d) = True Then
                        ComboBox2.DataSource = ArrayBlock3d
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'сельскохозяйственные угодья
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'заполняем вкладку с блоками
            Dim ArrayDeskBlk As String() = Nothing
            Dim NameTable As String = "Сельскохозяйственные угодья"
            Dim PutshTemplateFile As String = PutchBLK & ComboBox3.Text
            If File.Exists(PutshTemplateFile) = True Then
                'читаем все названия блоков с таблицы Растительность
                Dim boolDesk As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutshTemplateFile, NameTable, ArrayDeskBlk, "Блок")
                If IsArray(ArrayDeskBlk) = True Then
                    ComboBox2.DataSource = ArrayDeskBlk
                    boolSelectBlk = False
                End If
            End If
            Dim DeskBlk As String = ComboBox2.Text
            Dim ArrayNameBlk As String(,) = Nothing
            Dim boolName As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplateFile, DeskBlk, ArrayNameBlk, "Блок", NameTable)
            Dim NameImageZn As String = ""
            If IsArray(ArrayNameBlk) = True Then
                For i As Integer = 0 To ArrayNameBlk.GetUpperBound(1)
                    If ArrayNameBlk(0, i) Like "Отображение" Then
                        NameImageZn = ArrayNameBlk(1, i)
                    End If
                Next
            End If
            Dim PutchImage As String = PictureBox1.Tag
            If IsNothing(NameImageZn) = False Then
                If NameImageZn.Trim.Length > 0 Then
                    Try
                        Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & NameImageZn & ".png")
                        PictureBox2.Image = CType(MyImage, System.Drawing.Image)
                        PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                    Catch ex As FileNotFoundException
                    Catch ex As ArgumentException
                    Catch ex As IOException
                    End Try
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'меняем шаблон
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'заполняем вкладку с блоками
        Dim ArrayDeskBlk As String() = Nothing
        Dim NameTable As String = "Растительность"
        If RadioButton4.Checked = True Then
            NameTable = "Растительность"
        Else
            NameTable = "Сельскохозяйственные угодья"
        End If
        Dim PutshTemplateFile As String = PutchBLK & ComboBox3.Text
        If File.Exists(PutshTemplateFile) = True Then
            'читаем все названия блоков с таблицы Растительность
            Dim boolDesk As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutshTemplateFile, NameTable, ArrayDeskBlk, "Блок")
            If IsArray(ArrayDeskBlk) = True Then
                ComboBox2.DataSource = ArrayDeskBlk
                boolSelectBlk = False
            End If
        End If
        Dim DeskBlk As String = ComboBox2.Text
        Dim ArrayNameBlk As String(,) = Nothing
        Dim boolName As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplateFile, DeskBlk, ArrayNameBlk, "Блок", NameTable)
        Dim NameImageZn As String = ""
        If IsArray(ArrayNameBlk) = True Then
            For i As Integer = 0 To ArrayNameBlk.GetUpperBound(1)
                If ArrayNameBlk(0, i) Like "Отображение" Then
                    NameImageZn = ArrayNameBlk(1, i)
                End If
            Next
        End If
        Dim PutchImage As String = PictureBox1.Tag
        If IsNothing(NameImageZn) = False Then
            If NameImageZn.Trim.Length > 0 Then
                Try
                    Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & NameImageZn & ".png")
                    PictureBox2.Image = CType(MyImage, System.Drawing.Image)
                    PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                Catch ex As FileNotFoundException
                Catch ex As ArgumentException
                Catch ex As IOException
                End Try
            End If
        End If
    End Sub
    'выбор блока
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim PutshTemplateFile As String = PutchBLK & ComboBox3.Text
        Dim DeskBlk As String = ComboBox2.Text
        Dim NameTable As String = "Растительность"
        If RadioButton4.Checked = True Then
            NameTable = "Растительность"
        Else
            NameTable = "Сельскохозяйственные угодья"
        End If
        Dim ArrayNameBlk As String(,) = Nothing
        Dim boolName As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutshTemplateFile, DeskBlk, ArrayNameBlk, "Блок", NameTable)
        Dim NameImageZn As String = ""
        If IsArray(ArrayNameBlk) = True Then
            For i As Integer = 0 To ArrayNameBlk.GetUpperBound(1)
                If ArrayNameBlk(0, i) Like "Отображение" Then
                    NameImageZn = ArrayNameBlk(1, i)
                End If
            Next
        End If
        If IsNothing(NameImageZn) = False Then
            If NameImageZn.Trim.Length > 0 Then
                Try
                    Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & NameImageZn & ".jpg")
                    PictureBox2.Image = CType(MyImage, System.Drawing.Image)
                    PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
                Catch ex As FileNotFoundException
                Catch ex As ArgumentException
                Catch ex As IOException
                End Try
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        stbool = False
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'OK
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        stbool = True
    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Dim PutchImage As String = PictureBox1.Tag
        NumericUpDown2.Enabled = False
        Try
            Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & "Штриховка Сады.jpg")
            PictureBox1.Image = CType(MyImage, System.Drawing.Image)
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Catch ex As ArgumentException
        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Dim PutchImage As String = PictureBox1.Tag
        NumericUpDown2.Enabled = False
        Try
            Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & "Штриховка Сенокосы.jpg")
            PictureBox1.Image = CType(MyImage, System.Drawing.Image)
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Catch ex As ArgumentException
        End Try
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        Dim PutchImage As String = PictureBox1.Tag
        NumericUpDown2.Enabled = True
        Try
            Dim MyImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PutchImage & "Штриховка Лес.jpg")
            PictureBox1.Image = CType(MyImage, System.Drawing.Image)
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Catch ex As ArgumentException
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбрать блок
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim SelectBLK As PromptEntityResult = ed.GetEntity(vbLf & "Выберите образец блока для заполнения контура: ")
            If SelectBLK.Status = PromptStatus.Cancel Then
                Throw New System.Exception("Прервано пользователем")
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(SelectBLK.ObjectId, OpenMode.ForRead)
            If Not IsDBNull(acEnt) Then
                If TypeOf acEnt Is BlockReference Then
                    Dim BLK As BlockReference = acEnt
                    If BLK.IsDynamicBlock = True Then
                        Dim blockRef As BlockReference = acTrans.GetObject(SelectBLK.ObjectId, OpenMode.ForRead)
                        Dim blockDyn As BlockTableRecord = acTrans.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
                        Dim ArrayBlkName As String() = {blockDyn.Name}
                        ComboBox2.DataSource = ArrayBlkName
                    Else
                        Dim ArrayBlkName As String() = {BLK.Name}
                        ComboBox2.DataSource = ArrayBlkName
                    End If
                    boolSelectBlk = True
                    idBlk = BLK.Id
                Else
                    MsgBox("Вабранный объект не является блоком. Мокрос завершен")
                    Exit Sub
                End If
            End If
        End Using
    End Sub

    Private Sub HatchContour_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub
End Class