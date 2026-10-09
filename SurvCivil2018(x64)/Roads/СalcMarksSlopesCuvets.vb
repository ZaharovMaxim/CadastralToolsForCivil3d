Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Public Class СalcMarksSlopesCuvets
    Public boolIns As Boolean = False
    Public ActivCoridor As Autodesk.Civil.DatabaseServices.Corridor = Nothing
    Public ActivProfile As Autodesk.Civil.DatabaseServices.Profile = Nothing
    Public FeatureLine1 As FeatureLineCodeInfo = Nothing
    Public FeatureLine2 As FeatureLineCodeInfo = Nothing
    Public PatchFileTemplate As String = ""

    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
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
            Try
                For Each fileDWG As String In IO.Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = IO.Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        ReDim Preserve ArrayTemplates(CountArrayTemplates)
                        ArrayTemplates(CountArrayTemplates) = IO.Path.GetFileName(fileDWG)
                        If CountArrayTemplates = 0 Then
                            ComboBox4.Tag = PutchTables
                            PatchFileTemplate = fileDWG
                        End If
                        CountArrayTemplates += 1
                    End If
                Next
            Catch ex As io.IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
        If IsArray(ArrayTemplates) = True Then
            ComboBox4.DataSource = ArrayTemplates
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'открываем транзакцию
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'заполням список коридоров
            '1. Выбираем коридор
            Dim CollectionCoridor As CorridorCollection = CivilDoc.CorridorCollection
            Dim Coridor1 As Corridor = Nothing
            If CollectionCoridor.Count > 0 Then
                'добавляем в массив все коридоры
                Dim ArrayNamesCoridor As String() = Nothing 'массив для коридоров
                Dim CountArrayNamesCorridor As Integer = 0 'счетчик коридоров
                For Each id As ObjectId In CollectionCoridor
                    Coridor1 = tr.GetObject(id, OpenMode.ForRead) 'получаем доступ к корридору
                    If CountArrayNamesCorridor = 0 Then
                        ActivCoridor = Coridor1
                    End If
                    ReDim Preserve ArrayNamesCoridor(CountArrayNamesCorridor)
                    ArrayNamesCoridor(CountArrayNamesCorridor) = Coridor1.Name
                    CountArrayNamesCorridor += 1
                Next
                If IsArray(ArrayNamesCoridor) = True Then
                    ComboBox1.DataSource = ArrayNamesCoridor
                End If
                ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            End If
            '============================================================================================
            'заполняем список профилей
            Dim Profile1 As Autodesk.Civil.DatabaseServices.Profile = Nothing 'линия профиля по умолчанию
            Dim ArrayProfileName As String() = Nothing 'массив с профилями
            Dim CountArrayProfileName As Integer = 0 'счетчик профилей
            ReDim Preserve ArrayProfileName(CountArrayProfileName)
            ArrayProfileName(CountArrayProfileName) = ""
            CountArrayProfileName += 1
            Try
                Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                For Each id As ObjectId In btr
                    Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    If TypeOf ent Is Autodesk.Civil.DatabaseServices.Profile Then
                        Profile1 = ent
                        If CountArrayProfileName = 0 Then
                            ActivProfile = Profile1
                        End If
                        ReDim Preserve ArrayProfileName(CountArrayProfileName)
                        ArrayProfileName(CountArrayProfileName) = Profile1.Name
                        CountArrayProfileName += 1
                    End If
                Next
            Catch
            End Try
            If IsArray(ArrayProfileName) = True Then
                ComboBox6.DataSource = ArrayProfileName
            End If
            ComboBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим коды коридора
            '2. Находим коды
            If IsNothing(ActivCoridor) = False Then
                Dim ArrayCodeFL1 As String() = Nothing
                Dim ArrayCodeFL2 As String() = Nothing
                Dim CountArrayCodeFL As Integer = 0
                For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
                    Dim CodeNAme As String = featureLine.CodeName
                    ReDim Preserve ArrayCodeFL1(CountArrayCodeFL)
                    ArrayCodeFL1(CountArrayCodeFL) = CodeNAme
                    ReDim Preserve ArrayCodeFL2(CountArrayCodeFL)
                    ArrayCodeFL2(CountArrayCodeFL) = CodeNAme
                    CountArrayCodeFL += 1
                Next
                ComboBox2.DataSource = ArrayCodeFL1
                ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                ComboBox3.DataSource = ArrayCodeFL2
                ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            End If
            '=============================================
            'масштаб чертежа
            Dim ArrayScale As String() = Nothing
            Dim boolAnnoScale As Boolean = AcadFunction.FuncReadNameAnnoScaleDWG(ArrayScale)
            If IsArray(ArrayScale) = True Then
                ComboBox7.DataSource = ArrayScale
            End If
            ComboBox7.Text = acCurDb.Cannoscale.Name
            '=================================================================================================
            Dim colSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
            If colSurf.Count > 0 Then
                Dim ArraySurfName As String() = Nothing
                Dim CountArraySurfName As Integer = 0
                For i As Integer = 0 To colSurf.Count - 1
                    Dim Surf As TinSurface = tr.GetObject(colSurf.Item(i), OpenMode.ForRead)
                    ReDim Preserve ArraySurfName(CountArraySurfName)
                    ArraySurfName(CountArraySurfName) = Surf.Name
                    CountArraySurfName += 1
                Next i
                If IsArray(ArraySurfName) = True Then
                    ComboBox5.DataSource = ArraySurfName
                End If
            End If
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка отмена
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolIns = True
        If RadioButton4.Checked = True Then
            FeatureLine2 = Nothing
        End If
        If RadioButton3.Checked = True Then
            FeatureLine1 = Nothing
            FeatureLine2 = Nothing
        End If
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор коридора
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim ArrayCodeFL1 As String() = Nothing
        Dim ArrayCodeFL2 As String() = Nothing
        Dim CountArrayCodeFL As Integer = 0
        Dim NameCoridor As String = ComboBox1.SelectedValue 'имя корридора
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim CollectionCoridor As CorridorCollection = CivilDoc.CorridorCollection
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            For Each id As ObjectId In CollectionCoridor
                Dim Coridor1 As Corridor = Nothing
                Coridor1 = tr.GetObject(id, OpenMode.ForRead) 'получаем доступ к корридору
                If Coridor1.Name Like NameCoridor Then
                    ActivCoridor = Coridor1
                    Exit For
                End If
            Next
            For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
                Dim CodeNAme As String = featureLine.CodeName
                ReDim Preserve ArrayCodeFL1(CountArrayCodeFL)
                ArrayCodeFL1(CountArrayCodeFL) = CodeNAme
                ReDim Preserve ArrayCodeFL2(CountArrayCodeFL)
                ArrayCodeFL2(CountArrayCodeFL) = CodeNAme
                CountArrayCodeFL += 1
            Next
            ComboBox2.DataSource = ArrayCodeFL1
            ComboBox3.DataSource = ArrayCodeFL2
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбрана опция рассавлять метки по оси трассы
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            FeatureLine1 = Nothing
            FeatureLine2 = Nothing
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор кода 1
    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        Dim CodName1 As String = ComboBox3.SelectedValue
        For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
            Dim CodeNAme As String = featureLine.CodeName
            If CodeNAme Like CodName1 Then
                FeatureLine1 = featureLine
                Exit For
            End If
        Next
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор кода 2
    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        Dim CodName1 As String = ComboBox4.SelectedValue
        For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
            Dim CodeNAme As String = featureLine.CodeName
            If CodeNAme Like CodName1 Then
                FeatureLine2 = featureLine
                Exit For
            End If
        Next
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'расставить метки на харктерной линии коридора
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            Dim CodName1 As String = ComboBox2.Text
            For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
                Dim CodeNAme As String = featureLine.CodeName
                If CodeNAme Like CodName1 Then
                    FeatureLine1 = featureLine
                    Exit For
                End If
            Next
            ComboBox2.Enabled = True
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'расставить метки между характерными линиями
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then
            Dim CodName1 As String = ComboBox2.Text
            Dim CodName2 As String = ComboBox3.Text
            For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
                Dim CodeNAme As String = featureLine.CodeName
                If CodeNAme Like CodName1 Then
                    FeatureLine1 = featureLine
                    Exit For
                End If
            Next
            ComboBox2.Enabled = True
            For Each featureLine As FeatureLineCodeInfo In ActivCoridor.FeatureLineCodeInfos
                Dim CodeNAme As String = featureLine.CodeName
                If CodeNAme Like CodName2 Then
                    FeatureLine2 = featureLine
                    Exit For
                End If
            Next
            ComboBox3.Enabled = True
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            NumericUpDown3.Enabled = True
        Else
            NumericUpDown3.Enabled = False
        End If
    End Sub
    'меняем файл-шаблон
    Private Sub ComboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedValueChanged
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
        If IO.File.Exists(PutchTables & ComboBox4.Text) = True Then
            PatchFileTemplate = PutchTables & ComboBox4.Text
        End If
    End Sub

    Private Sub СalcMarksSlopesCuvets_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    End Sub
End Class