Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.Runtime
Imports System.Drawing.Graphics
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO

Public Class FormCreatePoint
    Public UstanovkiLoad As Integer = 0
    Public PutchTemlate As String = "d:\Civil3DToolsUtility\FileResources\Sample\"
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        '===============================================================================================================
        'закрепление точки
        Dim PointZakr As String() = {"Закрепление отсутствует", "Временный межевой знак", "Долговременный межевой знак"}
        ComboBox2.DataSource = PointZakr
        'метод определения координат
        Dim MetodCoord As String() = {"Геодезический метод", "Метод спутниковых геодезических измерений", "Аналитический метод", "Картометрический метод", "Фотограмметрический метод", "Иное описание"}
        ComboBox5.DataSource = MetodCoord
        'тип точки
        Dim PrefNumber As String() = {"", "т", "н", "%%u"}
        ComboBox4.DataSource = PrefNumber

        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName

        Dim ArrayFilesDirTemlates As String() = Nothing
        Dim CountPutchTemplateDir As Integer = 0
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemlate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In Directory.GetFiles(PutchTemlate)
                Try
                    Dim File1 As String = Path.GetFileName(fileDwg)
                    Dim ext As String = Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        ReDim Preserve ArrayFilesDirTemlates(CountPutchTemplateDir)
                        ArrayFilesDirTemlates(CountPutchTemplateDir) = File1
                        CountPutchTemplateDir += 1
                    End If
                Catch ex As ArgumentException
                End Try
            Next
            If IsArray(ArrayFilesDirTemlates) = True Then
                ComboBox3.DataSource = ArrayFilesDirTemlates
                ComboBox3.Tag = PutchTemlate
            End If
            Dim ArrayData As String() = Nothing
            Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
            If IO.File.Exists(PutchTemlateDwg) = True Then
                Dim boolDeskPoint As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Блок", "Проект планировки и межевания территории")
            End If
            If IsNothing(ArrayData) = True Then ArrayData = {""}
            If IsArray(ArrayData) = False Then ArrayData = {""}
            ComboBox1.DataSource = ArrayData
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ОК
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        UstanovkiLoad = 1
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'отмена
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        UstanovkiLoad = 0
        Me.Hide()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'масштабы
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        NumericUpDown1.Value = 10000
    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        NumericUpDown1.Value = 7000
    End Sub
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        NumericUpDown1.Value = 5000
    End Sub
    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        NumericUpDown1.Value = 2000
    End Sub
    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        NumericUpDown1.Value = 1000
    End Sub
    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        NumericUpDown1.Value = 500
    End Sub
    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        NumericUpDown1.Value = 100
    End Sub
    Private Sub RadioButton18_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton18.CheckedChanged
        NumericUpDown1.Value = 3000
    End Sub
    Private Sub RadioButton17_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton17.CheckedChanged
        NumericUpDown1.Value = 200
    End Sub
    'земельные участки образуемые
    Private Sub RadioButton10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton10.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton11.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton12.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton13.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Метка", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
    End Sub
    'здание (строение) сооружение
    Private Sub RadioButton9_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton11.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton12.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton13.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "Метка", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
    End Sub
    'сооружения линейного типа
    Private Sub RadioButton22_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton22.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton11.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton12.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton13.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "Метка", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
    End Sub
    'территориальные зоны
    Private Sub RadioButton20_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton20.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton11.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton12.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton13.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "Метка", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
    End Sub
    'красные линии
    Private Sub RadioButton21_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton21.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton11.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton12.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton13.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "Метка", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
    End Sub
    'блок Автокад
    Private Sub RadioButton11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton11.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton10.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton9.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton22.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton20.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "Блок", "Проект планировки и межевания территории")
            ElseIf RadioButton23.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Блок", "Благоустройство и озеленение территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "Блок", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
        RadioButton19.Enabled = False
        NumericUpDown3.Value = 1.5
    End Sub
    'точка COGO
    Private Sub RadioButton12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton12.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton10.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton9.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton22.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton20.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            ElseIf RadioButton23.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Блок", "Благоустройство и озеленение территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "ТочкаCOGO", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
        RadioButton19.Enabled = False
        NumericUpDown3.Value = 1.5
    End Sub
    'мультивыноска
    Private Sub RadioButton13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton13.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton10.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            ElseIf RadioButton9.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            ElseIf RadioButton22.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            ElseIf RadioButton20.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            ElseIf RadioButton23.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Блок", "Благоустройство и озеленение территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "Мультивыноска", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
        RadioButton19.Enabled = True
        NumericUpDown3.Value = 5
    End Sub
    'метка
    Private Sub RadioButton16_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton16.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton10.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Земельные участки образуемые", ArrayData, "Метка", "Проект планировки и межевания территории")
            ElseIf RadioButton9.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Здания (строения) образуемые", ArrayData, "Метка", "Проект планировки и межевания территории")
            ElseIf RadioButton22.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Объекты капитального строительства (инженерные коммуникации)", ArrayData, "Метка", "Проект планировки и межевания территории")
            ElseIf RadioButton20.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Зоны с особыми условиями использования территорий", ArrayData, "Метка", "Проект планировки и межевания территории")
            ElseIf RadioButton23.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Блок", "Благоустройство и озеленение территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Линии градостроительного регулирования", ArrayData, "Метка", "Проект планировки и межевания территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
        RadioButton19.Enabled = True
        NumericUpDown3.Value = 5
    End Sub
    '==========================================================================================================
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub ComboBox1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ComboBox1.DrawItem
        Dim Text As String = Me.ComboBox1.GetItemText(ComboBox1.Items(e.Index))
        e.DrawBackground()
        Dim br As Drawing.SolidBrush = New Drawing.SolidBrush(e.ForeColor)
        Using br
            e.Graphics.DrawString(Text, e.Font, br, e.Bounds)
            If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
                If (Text.Length > Math.Round(Convert.ToDecimal(Me.ComboBox1.Size.Width / 8))) Then
                    Me.ToolTip1.Show(Text, ComboBox1, e.Bounds.Right, e.Bounds.Bottom)
                End If
            Else
                Me.ToolTip1.Hide(ComboBox1)
            End If
            e.DrawFocusRectangle()
        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_MouseHover(sender As Object, e As EventArgs) Handles ComboBox1.MouseHover
        ComboBox1.DrawMode = DrawMode.OwnerDrawVariable
    End Sub

    Private Sub FormCreatePoint_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
    'лесной участок
    Private Sub RadioButton23_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton23.CheckedChanged
        Dim PutchTemlateDwg As String = PutchTemlate & ComboBox3.Text
        Dim ArrayData As String() = Nothing
        If IO.File.Exists(PutchTemlateDwg) = True Then
            If RadioButton11.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Блок", "Благоустройство и озеленение территории")
            ElseIf RadioButton12.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "ТочкаCOGO", "Благоустройство и озеленение территории")
            ElseIf RadioButton13.Checked = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Мультивыноска", "Благоустройство и озеленение территории")
            Else
                Dim boolFindObj As Boolean = AcadFunction.FuncReadDeskriptionObjectAllDestDWG(PutchTemlateDwg, "Снос зеленых насаждений", ArrayData, "Метка", "Благоустройство и озеленение территории")
            End If
        End If
        If IsNothing(ArrayData) = True Then ArrayData = {""}
        If IsArray(ArrayData) = False Then ArrayData = {""}
        ComboBox1.DataSource = ArrayData
    End Sub
End Class