Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports System.Windows.Forms
Imports System.IO

Public Class PlanStatement
    '==========================================================================
    'глобальные переменные для передачи их в основно тело программы
    Public boolCoridor As Boolean = False
    Public PKStart As Double = 0
    Public PKEnd As Double = 0
    Public ArrayCodes As String() = Nothing
    Public NameFeatureLine1 As String = ""
    Public NameFeatureLine2 As String = ""
    Public CollectionCoridor As CorridorCollection = Nothing
    Dim DbCorridorCollection As DBObjectCollection = Nothing
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        'находим все директории с файлами Excel
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        Dim PutchTables As String = "d:\AutodeskCivilTools\TemplateTables\Roads\Ведомости планировки\Пользовательские\" 'пути для шаблона к образцам блоков
        If IO.Directory.Exists(parentPatch2 & "\TemplateTables\Roads\Ведомости планировки\Пользовательские\") Then
            PutchTables = parentPatch2 & "\TemplateTables\Roads\Ведомости планировки\Пользовательские\"
        End If
        ComboBox7.Tag = PutchTables
        Dim ArrayTemplates As String() = Nothing
        Dim CountArrayTemplates As Integer = 0
        'читаем шаблоны условных знаков
        If IO.Directory.Exists(PutchTables) = True Then
            Try
                For Each fileExcel As String In Directory.GetFiles(PutchTables)
                    'читаем расширение
                    Dim ext As String = Path.GetExtension(fileExcel)
                    If ext Like ".xlsx" Then
                        ReDim Preserve ArrayTemplates(CountArrayTemplates)
                        ArrayTemplates(CountArrayTemplates) = Path.GetFileName(fileExcel)
                        CountArrayTemplates += 1
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
        If IsArray(ArrayTemplates) = True Then
            ComboBox7.DataSource = ArrayTemplates
        End If
        ComboBox7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование

        CollectionCoridor = CivilDoc.CorridorCollection
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Dim Coridor As Corridor = Nothing
        Dim bLine As Baseline = Nothing

        Dim ArrayNamesCoridor As String() = Nothing
        Dim CountArrayNamesCorridor As Integer = 0
        Dim boolCor As Boolean = False

        Dim ArrayBLines As String() = Nothing
        Using tr
            'добавляем в массив все коридоры
            If CollectionCoridor.Count > 0 Then
                For Each id As ObjectId In CollectionCoridor
                    Dim Coridor1 As Corridor = tr.GetObject(id, OpenMode.ForRead) 'получаем доступ к корридору
                    If boolCor = False Then
                        Coridor = Coridor1
                        boolCor = True
                    End If
                    ReDim Preserve ArrayNamesCoridor(CountArrayNamesCorridor)
                    ArrayNamesCoridor(CountArrayNamesCorridor) = Coridor1.Name
                    CountArrayNamesCorridor += 1
                Next
                ComboBox1.DataSource = ArrayNamesCoridor
                ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                If IsNothing(Coridor) = False Then
                    'получаем коды корридора
                    Dim CountArrayCode As Integer = 0
                    Dim PtColl2 As FeatureLineCodeInfoCollection = Coridor.FeatureLineCodeInfos
                    If PtColl2.Count > 0 Then
                        Dim ArrayCode1 As String() = Nothing 'массив для кода1
                        Dim ArrayCode2 As String() = Nothing 'массив для кода2
                        For Each cd As FeatureLineCodeInfo In PtColl2
                            Dim CodeN As String = cd.CodeName
                            ReDim Preserve ArrayCode1(CountArrayCode)
                            ReDim Preserve ArrayCode2(CountArrayCode)
                            ArrayCode1(CountArrayCode) = CodeN
                            ArrayCode2(CountArrayCode) = CodeN
                            CountArrayCode += 1
                        Next
                        If CountArrayTemplates > 0 Then
                            ReDim Preserve ArrayCode1(CountArrayCode)
                            ReDim Preserve ArrayCode2(CountArrayCode)
                            ArrayCode1(CountArrayCode) = "Выбрать все"
                            ArrayCode2(CountArrayCode) = "Выбрать все"
                        End If
                        ComboBox3.DataSource = ArrayCode1
                        ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                        If ComboBox7.Text.Trim.Length > 0 Then
                            ComboBox3.Text = "Выбрать все"
                        End If
                        ComboBox4.DataSource = ArrayCode2
                        ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                        If ComboBox7.Text.Trim.Length > 0 Then
                            ComboBox4.Text = "Выбрать все"
                        End If
                    End If
                    'получаем коллекцию базовых линий
                    Dim blineCollection As BaselineCollection = Coridor.Baselines 'получаем доступ к базовой линии
                    If blineCollection.Count > 0 Then
                        Dim CountBLines As Integer = 0
                        For Each TempbLine As Baseline In blineCollection
                            Dim NameBLine As String = TempbLine.Name
                            ReDim Preserve ArrayBLines(CountBLines)
                            ArrayBLines(CountBLines) = NameBLine
                            CountBLines += 1
                            If CountBLines = 1 Then
                                bLine = TempbLine
                            End If
                        Next
                        ComboBox2.DataSource = ArrayBLines
                        ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                    End If
                    'начальный и конечный пикеты
                    If IsNothing(bLine) = False Then
                        Dim StartSt As Double = bLine.StartStation
                        Dim EndSt As Double = bLine.EndStation
                        TextBox1.Text = FuncFormatZN.FuncFormatPK(StartSt)
                        TextBox2.Text = FuncFormatZN.FuncFormatPK(EndSt)
                        PKStart = StartSt
                        PKEnd = EndSt
                    End If
                    'получаем коллекцию груп самплов
                    Try
                        Dim AlignID As ObjectId = bLine.AlignmentId
                        Dim ObjAlign As Alignment = tr.GetObject(AlignID, OpenMode.ForRead) 'получаем доступ к трассе
                        Dim SampleLineGroupCollection As ObjectIdCollection = ObjAlign.GetSampleLineGroupIds
                        Dim ArraySampleGroup As String() = Nothing
                        Dim CountArraySampleGroup As Integer = 0
                        For Each sampleLineId As ObjectId In SampleLineGroupCollection
                            Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                            ReDim Preserve ArraySampleGroup(CountArraySampleGroup)
                            ArraySampleGroup(CountArraySampleGroup) = sampleGroupLine.Name
                            CountArraySampleGroup += 1
                        Next
                        ComboBox5.DataSource = ArraySampleGroup
                        ComboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                    Catch ex As NullReferenceException
                    End Try
                End If
            End If
        End Using
        Dim ArrayStrPK As String() = {"ПК+", "0.00"}
        ComboBox6.DataSource = ArrayStrPK
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка ОТМЕНА
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolCoridor = False
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка OK
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolCoridor = True
        NameFeatureLine1 = ComboBox3.Text
        NameFeatureLine2 = ComboBox4.Text
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub UstanovkiCoridor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ComboBox3.DrawMode = DrawMode.OwnerDrawFixed
        Me.ComboBox4.DrawMode = DrawMode.OwnerDrawFixed
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор коридора из списка
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim NameCoridor As String = ComboBox1.Text
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        CollectionCoridor = CivilDoc.CorridorCollection
        Dim bline As Baseline = Nothing
        If CollectionCoridor.Count < 1 Then Exit Sub
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim IDCoridor As Autodesk.AutoCAD.DatabaseServices.ObjectId = CollectionCoridor.Item(NameCoridor)
            Dim Coridor As Corridor = tr.GetObject(IDCoridor, OpenMode.ForRead) 'получаем доступ к корридору
            'получаем коллекцию базовых линий
            Dim blineCollection As BaselineCollection = Coridor.Baselines 'получаем доступ к базовой линии
            If blineCollection.Count > 0 Then
                Dim CountBLines As Integer = 0
                Dim ArrayBLines As String() = Nothing
                For Each tempbline As Baseline In blineCollection
                    Dim NameBLine As String = tempbline.Name
                    ReDim Preserve ArrayBLines(CountBLines)
                    ArrayBLines(CountBLines) = NameBLine
                    CountBLines += 1
                    If CountBLines = 1 Then
                        bline = tempbline
                    End If
                Next
                ComboBox2.DataSource = ArrayBLines
                ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                'заполняем пикетаж 
                If IsNothing(bline) = False Then
                    Dim StartSt As Double = bline.StartStation
                    Dim EndSt As Double = bline.EndStation
                    TextBox1.Text = FuncFormatZN.FuncFormatPK(StartSt)
                    TextBox2.Text = FuncFormatZN.FuncFormatPK(EndSt)
                    PKStart = StartSt
                    PKEnd = EndSt
                End If
            End If
            'заполняем коды
            If IsNothing(Coridor) = False Then
                'получаем коды корридора
                Dim CountArrayCode As Integer = 0
                Dim PtColl2 As FeatureLineCodeInfoCollection = Coridor.FeatureLineCodeInfos
                If PtColl2.Count > 0 Then
                    Dim ArrayCode1 As String() = Nothing 'массив для кода1
                    Dim ArrayCode2 As String() = Nothing 'массив для кода2
                    For Each cd As FeatureLineCodeInfo In PtColl2
                        Dim CodeN As String = cd.CodeName
                        ReDim Preserve ArrayCode1(CountArrayCode)
                        ReDim Preserve ArrayCode2(CountArrayCode)
                        ArrayCode1(CountArrayCode) = CodeN
                        ArrayCode2(CountArrayCode) = CodeN
                        CountArrayCode += 1
                    Next
                    If ComboBox7.Text.Trim.Length > 0 Then
                        ReDim Preserve ArrayCode1(CountArrayCode)
                        ReDim Preserve ArrayCode2(CountArrayCode)
                        ArrayCode1(CountArrayCode) = "Выбрать все"
                        ArrayCode2(CountArrayCode) = "Выбрать все"
                    End If
                    ComboBox3.DataSource = ArrayCode1
                    ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                    If ComboBox7.Text.Trim.Length > 0 Then
                        ComboBox3.Text = "Выбрать все"
                    End If
                    ComboBox4.DataSource = ArrayCode2
                    ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                    If ComboBox7.Text.Trim.Length > 0 Then
                        ComboBox4.Text = "Выбрать все"
                    End If
                End If
            End If
            'получаем коллекцию груп самплов
            Try
                Dim AlignID As ObjectId = bline.AlignmentId
                Dim ObjAlign As Alignment = tr.GetObject(AlignID, OpenMode.ForRead) 'получаем доступ к трассе
                Dim SampleLineGroupCollection As ObjectIdCollection = ObjAlign.GetSampleLineGroupIds
                Dim ArraySampleGroup As String() = Nothing
                Dim CountArraySampleGroup As Integer = 0
                For Each sampleLineId As ObjectId In SampleLineGroupCollection
                    Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                    ReDim Preserve ArraySampleGroup(CountArraySampleGroup)
                    ArraySampleGroup(CountArraySampleGroup) = sampleGroupLine.Name
                    CountArraySampleGroup += 1
                Next
                ComboBox5.DataSource = ArraySampleGroup
                ComboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            Catch ex As NullReferenceException
            End Try
        End Using
        locker.Dispose()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка начальный пикет
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim PK As String = TextBox1.Text
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim NameCoridor As String = ComboBox1.Text.Trim
        '2. Получаем базовую линию
        Dim NameBaseLine As String = ComboBox2.Text.Trim
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim IdCoridor As ObjectId = CivilDoc.CorridorCollection.Item(NameCoridor)
            Dim ObjCoridor As Corridor = tr.GetObject(IdCoridor, OpenMode.ForRead) 'получаем доступ к корридору
            Dim ObjBaseLine As Baseline = ObjCoridor.Baselines.Item(NameBaseLine) 'получаем доступ к базовой линии
            Dim AlignId As ObjectId = ObjBaseLine.AlignmentId
            Dim ObjAlign As Alignment = tr.GetObject(AlignId, OpenMode.ForRead) 'получаем доступ к трассе
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку начала трассы: "
            Dim pPtRes As PromptPointResult = ed.GetPoint(pPtOpts)
            Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
            Dim PKVedStartAlign As Double = 0
            Dim OfsetVedStartAlign As Double = 0
            If pPtRes.Status = PromptStatus.None Then
                TextBox1.Text = PK
            Else
                Try
                    ObjAlign.StationOffset(ptStart.X, ptStart.Y, PKVedStartAlign, OfsetVedStartAlign)
                    TextBox1.Text = FuncFormatZN.FuncFormatPK(PKVedStartAlign)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                    TextBox1.Text = PK
                    Exit Sub
                End Try
            End If
        End Using
        locker.Dispose()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'кнопка конечный пикет
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
        Dim PK As String = TextBox2.Text
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim NameCoridor As String = ComboBox1.Text.Trim
        '2. Получаем базовую линию
        Dim NameBaseLine As String = ComboBox2.Text.Trim
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim IdCoridor As ObjectId = CivilDoc.CorridorCollection.Item(NameCoridor)
            Dim ObjCoridor As Corridor = tr.GetObject(IdCoridor, OpenMode.ForRead) 'получаем доступ к корридору
            Dim ObjBaseLine As Baseline = ObjCoridor.Baselines.Item(NameBaseLine) 'получаем доступ к базовой линии
            Dim AlignId As ObjectId = ObjBaseLine.AlignmentId
            Dim ObjAlign As Alignment = tr.GetObject(AlignId, OpenMode.ForRead) 'получаем доступ к трассе
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            pPtOpts.AllowNone = True
            pPtOpts.Message = vbLf & "Укажите точку конца трассы: "
            Dim pPtRes As PromptPointResult = ed.GetPoint(pPtOpts)
            Dim ptStart As Point3d = pPtRes.Value 'координаты точки начала трассы
            Dim PKVedStartAlign As Double = 0
            Dim OfsetVedStartAlign As Double = 0
            If pPtRes.Status = PromptStatus.None Then
                TextBox2.Text = PK
            Else
                Try
                    ObjAlign.StationOffset(ptStart.X, ptStart.Y, PKVedStartAlign, OfsetVedStartAlign)
                    TextBox2.Text = FuncFormatZN.FuncFormatPK(PKVedStartAlign)
                Catch ex As Autodesk.Civil.PointNotOnEntityException
                    MsgBox("Не верно указана точка начала!!!")
                    TextBox2.Text = PK
                    Exit Sub
                End Try
            End If
        End Using
        locker.Dispose()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор базовой линии из списка
    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        Dim NameCoridor As String = ComboBox1.Text
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        CollectionCoridor = CivilDoc.CorridorCollection
        Dim bline As Baseline = Nothing
        If CollectionCoridor.Count < 1 Then Exit Sub
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim IDCoridor As Autodesk.AutoCAD.DatabaseServices.ObjectId = CollectionCoridor.Item(NameCoridor)
            Dim Coridor As Corridor = tr.GetObject(IDCoridor, OpenMode.ForRead) 'получаем доступ к корридору
            'получаем коллекцию базовых линий
            Dim blineCollection As BaselineCollection = Coridor.Baselines 'получаем доступ к базовой линии
            If blineCollection.Count > 0 Then
                For Each bline In blineCollection
                    Dim NameBLine As String = bline.Name
                    If NameBLine Like ComboBox2.Text Then
                        Exit For
                    End If
                Next
                'заполняем пикетаж 
                If IsNothing(bline) = False Then
                    Dim StartSt As Double = bline.StartStation
                    Dim EndSt As Double = bline.EndStation
                    TextBox1.Text = FuncFormatZN.FuncFormatPK(StartSt)
                    TextBox2.Text = FuncFormatZN.FuncFormatPK(EndSt)
                    PKStart = StartSt
                    PKEnd = EndSt
                End If
            End If
            'заполняем коды
            If IsNothing(Coridor) = False Then
                'получаем коды корридора
                Dim CountArrayCode As Integer = 0
                Dim PtColl2 As FeatureLineCodeInfoCollection = Coridor.FeatureLineCodeInfos
                If PtColl2.Count > 0 Then
                    Dim ArrayCode1 As String() = Nothing 'массив для кода1
                    Dim ArrayCode2 As String() = Nothing 'массив для кода2
                    For Each cd As FeatureLineCodeInfo In PtColl2
                        Dim CodeN As String = cd.CodeName
                        ReDim Preserve ArrayCode1(CountArrayCode)
                        ReDim Preserve ArrayCode2(CountArrayCode)
                        ArrayCode1(CountArrayCode) = CodeN
                        ArrayCode2(CountArrayCode) = CodeN
                        CountArrayCode += 1
                    Next
                    ReDim Preserve ArrayCode1(CountArrayCode)
                    ReDim Preserve ArrayCode2(CountArrayCode)
                    ArrayCode1(CountArrayCode) = "Выбрать все"
                    ArrayCode2(CountArrayCode) = "Выбрать все"
                    ComboBox3.DataSource = ArrayCode1
                    ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                    ComboBox3.Text = "Выбрать все"
                    ComboBox4.DataSource = ArrayCode2
                    ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
                    ComboBox4.Text = "Выбрать все"
                End If
            End If
            'получаем коллекцию груп самплов
            Try
                Dim AlignID As ObjectId = bline.AlignmentId
                Dim ObjAlign As Alignment = tr.GetObject(AlignID, OpenMode.ForRead) 'получаем доступ к трассе
                Dim SampleLineGroupCollection As ObjectIdCollection = ObjAlign.GetSampleLineGroupIds
                Dim ArraySampleGroup As String() = Nothing
                Dim CountArraySampleGroup As Integer = 0
                For Each sampleLineId As ObjectId In SampleLineGroupCollection
                    Dim sampleGroupLine As SampleLineGroup = tr.GetObject(sampleLineId, OpenMode.ForRead)
                    ReDim Preserve ArraySampleGroup(CountArraySampleGroup)
                    ArraySampleGroup(CountArraySampleGroup) = sampleGroupLine.Name
                    CountArraySampleGroup += 1
                Next
                ComboBox5.DataSource = ArraySampleGroup
                ComboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList 'запрет на редактирование
            Catch ex As NullReferenceException
            End Try
        End Using
        locker.Dispose()
    End Sub

    Private Sub ComboBox3_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ComboBox3.DrawItem
        Dim Text As String = Me.ComboBox3.GetItemText(ComboBox3.Items(e.Index))
        e.DrawBackground()
        Dim br As Drawing.SolidBrush = New Drawing.SolidBrush(e.ForeColor)
        Using br
            e.Graphics.DrawString(Text, e.Font, br, e.Bounds)
            If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
                If (Text.Length > Math.Round(Convert.ToDecimal(Me.ComboBox3.Size.Width / 8))) Then
                    Me.ToolTip1.Show(Text, ComboBox3, e.Bounds.Right, e.Bounds.Bottom)
                End If
            Else
                Me.ToolTip1.Hide(ComboBox3)
            End If
            e.DrawFocusRectangle()
        End Using
    End Sub

    Private Sub ComboBox4_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ComboBox4.DrawItem
        Dim Text As String = Me.ComboBox4.GetItemText(ComboBox3.Items(e.Index))
        e.DrawBackground()
        Dim br As Drawing.SolidBrush = New Drawing.SolidBrush(e.ForeColor)
        Using br
            e.Graphics.DrawString(Text, e.Font, br, e.Bounds)
            If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
                If (Text.Length > Math.Round(Convert.ToDecimal(Me.ComboBox4.Size.Width / 8))) Then
                    Me.ToolTip1.Show(Text, ComboBox4, e.Bounds.Right, e.Bounds.Bottom)
                End If
            Else
                Me.ToolTip1.Hide(ComboBox4)
            End If
            e.DrawFocusRectangle()
        End Using
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

    End Sub
End Class