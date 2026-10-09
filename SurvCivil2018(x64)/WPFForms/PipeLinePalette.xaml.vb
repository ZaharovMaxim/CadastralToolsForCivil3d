Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports System.IO
Imports System.Drawing 'для рисунков jpg, tiff
Imports System.Collections.ObjectModel
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows
Imports Microsoft.Office.Interop

Public Class PipeLinePalette
    Public PipePalette As Autodesk.AutoCAD.Windows.PaletteSet = Nothing
    Public indCount As Integer = 0
    Public Shared custdata As ObservableCollection(Of Customer) = New ObservableCollection(Of Customer)
    Public Sub New()
        InitializeComponent()
        '========================================================================================
        'запускаем реактор
        '' Get the current document
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        AddHandler acDoc.ImpliedSelectionChanged, AddressOf FuncSelectionPipe

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'столбцы
    Public Class Customer
        Public Property Поз() As Integer
        Public Property В_отчет() As Boolean
        Public Property Имя_свойства() As String
        Public Property Описание_свойства() As String
        Public Property Значение() As String
        'Public Property Status() As OrderStatus
    End Class
    Public Enum OrderStatus
        None
        [New]
        Processing
        Shipped
        Received
    End Enum
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функции
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Public Function FuncReadPropertiesPipe(ByVal obj As Autodesk.AutoCAD.DatabaseServices.Entity, ByVal tr As Transaction, Optional ByVal ArrayField As String() = Nothing) As Boolean
        FuncReadPropertiesPipe = True
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем поля труб и колодцев
        If TypeOf obj Is [Structure] Then
            '=========================================================================
            'меняем название кнопки
            custdata.Clear()
            Dim TextBlock1 As TextBlock = New TextBlock
            TextBlock1.Text = "Колодец"
            TextBlock1.FontSize = 16
            TextBlock1.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            TextBlock1.VerticalAlignment = System.Windows.VerticalAlignment.Center
            button1.Content = TextBlock1
            '===========================================================================
            Dim StructObj As [Structure] = obj 'делаем ссылку на колодец
            Dim netWorkStructure As String = StructObj.NetworkName 'имя трубопроводной сети
            textBox.Text = netWorkStructure
            textBox.Tag = StructObj.Handle.ToString
            Dim boolField As Boolean = False
            'имя колодца 0
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Name" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Name"
                    Acustim.Описание_свойства = "Имя элемента"
                    Acustim.Значение = StructObj.Name
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Name"
                    Acustim.Описание_свойства = "Имя элемента"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '1
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Desk" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Desk"
                    Acustim.Описание_свойства = "Описание элемента"
                    Acustim.Значение = StructObj.Description
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Desk"
                    Acustim.Описание_свойства = "Описание элемента"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '2
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "RefAlignmentName" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RefAlignmentName"
                    Acustim.Описание_свойства = "Имя трассы"
                    Acustim.Значение = StructObj.RefAlignmentName
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RefAlignmentName"
                    Acustim.Описание_свойства = "Имя трассы"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '3
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "RefSurfaceName" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RefSurfaceName"
                    Acustim.Описание_свойства = "Имя поверхности"
                    Acustim.Значение = StructObj.RefSurfaceName
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RefSurfaceName"
                    Acustim.Описание_свойства = "Имя поверхности"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '4
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Station" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Station"
                    Acustim.Описание_свойства = "Пикет"
                    Dim PK As Double = StructObj.Station
                    Dim intPK As Integer = Int(PK / 100)
                    Dim DblPk As Double = PK - intPK * 100
                    Dim DblPKFormat As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(DblPk, 2, 3)
                    Acustim.Значение = intPK & "+" & DblPKFormat
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As CivilException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Station"
                    Acustim.Описание_свойства = "Пикет"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '5
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Offset" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Offset"
                    Acustim.Описание_свойства = "Смещение"
                    Dim ofsetPk As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Offset, 2, 3)
                    If ofsetPk < 0 Then
                        Acustim.Значение = "Лево " & Math.Abs(ofsetPk).ToString
                    ElseIf ofsetPk > 0 Then
                        Acustim.Значение = "Право " & Math.Abs(ofsetPk).ToString
                    Else
                        Acustim.Значение = 0
                    End If
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As CivilException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Offset"
                    Acustim.Описание_свойства = "Смещение"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '6
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Northing" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Northing"
                    Acustim.Описание_свойства = "Северное положение"
                    Acustim.Значение = Math.Round(StructObj.Northing, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '7
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Easting" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Easting"
                    Acustim.Описание_свойства = "Восточное положение"
                    Acustim.Значение = Math.Round(StructObj.Easting, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '8
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "ElevTop" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ElevTop"
                    Acustim.Описание_свойства = "Отметка верха колодца"
                    Acustim.Значение = Math.Round(StructObj.RimElevation, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ElevTop"
                    Acustim.Описание_свойства = "Отметка верха колодца"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '9
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "ElevDown" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ElevDown"
                    Acustim.Описание_свойства = "Отметка дна колодца"
                    Acustim.Значение = Math.Round(StructObj.SumpElevation, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ElevDown"
                    Acustim.Описание_свойства = "Отметка дна колодца"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '10
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "PartDesk" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartDesk"
                    Acustim.Описание_свойства = "Название элемента"
                    Acustim.Значение = StructObj.PartDescription 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartDesk"
                    Acustim.Описание_свойства = "Название элемента"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '11
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "PartSizeName" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartSizeName"
                    Acustim.Описание_свойства = "Название размера элемента"
                    Acustim.Значение = StructObj.PartSizeName 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartSizeName"
                    Acustim.Описание_свойства = "Название размера элемента"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '12
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "PartSubType" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartSubType"
                    Acustim.Описание_свойства = "Описание подтипа"
                    Acustim.Значение = StructObj.PartSubType 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartSubType"
                    Acustim.Описание_свойства = "Описание подтипа"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '13
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Rotation" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Rotation"
                    Acustim.Описание_свойства = "Разворот колодца"
                    Dim an As Autodesk.AutoCAD.Runtime.AngularUnitFormat = Autodesk.AutoCAD.Runtime.AngularUnitFormat.DegreesMinutesSeconds
                    Acustim.Значение = Autodesk.AutoCAD.Runtime.Converter.AngleToString(StructObj.Rotation, an, 6)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Rotation"
                    Acustim.Описание_свойства = "Разворот колодца"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '14
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "SumpDepth" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "SumpDepth"
                    Acustim.Описание_свойства = "Глубина отстойника"
                    Acustim.Значение = Math.Round(StructObj.SumpDepth, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "SumpDepth"
                    Acustim.Описание_свойства = "Глубина отстойника"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '15
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "RimElevation" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RimElevation"
                    Acustim.Описание_свойства = "Отметка обечайки"
                    Acustim.Значение = StructObj.RimElevation
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RimElevation"
                    Acustim.Описание_свойства = "Отметка обечайки"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '16
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Diameter" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Diameter"
                    Acustim.Описание_свойства = "Диаметр колодца, мм"
                    Acustim.Значение = Math.Round(StructObj.DiameterOrWidth * 1000, 0) 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Diameter"
                    Acustim.Описание_свойства = "Диаметр колодца, мм"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '17
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Lenght" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Lenght"
                    Acustim.Описание_свойства = "Длина колодца, мм"
                    Acustim.Значение = Math.Round(StructObj.Length * 1000, 0) 'только чтениe
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Lenght"
                    Acustim.Описание_свойства = "Длина колодца, мм"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '18
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Height" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Height"
                    Acustim.Описание_свойства = "Высота колодца"
                    Acustim.Значение = Math.Round(StructObj.Height, 0)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Height"
                    Acustim.Описание_свойства = "Высота колодца"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '19
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "RimToSumpHeight" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RimToSumpHeight"
                    Acustim.Описание_свойства = "RimToSumpHeight, м"
                    Acustim.Значение = StructObj.RimToSumpHeight
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RimToSumpHeight"
                    Acustim.Описание_свойства = "RimToSumpHeight, м"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '20
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "BarrelPipeClearance" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "BarrelPipeClearance"
                    Acustim.Описание_свойства = "Зазор трубы"
                    Acustim.Значение = StructObj.BarrelPipeClearance
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '21
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "SurfElevAtInsPoint" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "SurfElevAtInsPoint"
                    Acustim.Описание_свойства = "Отметка земли, м"
                    Acustim.Значение = Math.Round(StructObj.SurfaceElevationAtInsertionPoint, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As CivilException
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "SurfElevAtInsPoint"
                    Acustim.Описание_свойства = "Отметка земли, м"
                    Acustim.Значение = "Не определено"
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                End Try
            End If
            boolField = False
            '22
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "ConeHeight" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ConeHeight"
                    Acustim.Описание_свойства = "Размер конуса"
                    Acustim.Значение = StructObj.ConeHeight
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '23
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "ConnectedPartCount" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ConnectedPartCount"
                    Acustim.Описание_свойства = "Количество типов труб"
                    Acustim.Значение = StructObj.ConnectedPartCount
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '24
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "ConnectedPipesCount" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ConnectedPipesCount"
                    Acustim.Описание_свойства = "Количество труб"
                    Acustim.Значение = StructObj.ConnectedPipesCount
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '25
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "SurfaceAdjustmentValue" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "SurfaceAdjustmentValue"
                    Acustim.Описание_свойства = "Отметка поверхности корректированная"
                    Acustim.Значение = StructObj.SurfaceAdjustmentValue
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '26
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "FloorThickness" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "FloorThickness"
                    Acustim.Описание_свойства = "Толщина нижней части"
                    Acustim.Значение = StructObj.FloorThickness
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '27
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "RuleSetStyleName" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RuleSetStyleName"
                    Acustim.Описание_свойства = "Стиль колодца"
                    Acustim.Значение = StructObj.RuleSetStyleName
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '28
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "FrameDiameter" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "FrameDiameter"
                    Acustim.Описание_свойства = "Диаметр рамы колодца"
                    Acustim.Значение = StructObj.FrameDiameter
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '29
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "FrameHeight" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "FrameHeight"
                    Acustim.Описание_свойства = "Высота рамы колодца"
                    Acustim.Значение = StructObj.FrameHeight
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '30
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Grate" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Grate"
                    Acustim.Описание_свойства = "Решетка"
                    Acustim.Значение = StructObj.Grate
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '31
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "HeadwallBaseThickness" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "HeadwallBaseThickness"
                    Acustim.Описание_свойства = "Толщина оголовка"
                    Acustim.Значение = StructObj.HeadwallBaseThickness
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '32
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "HeadwallBaseWidth" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "HeadwallBaseWidth"
                    Acustim.Описание_свойства = "Высота оголовка"
                    Acustim.Значение = StructObj.HeadwallBaseWidth
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '33
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "InnerDiameterOrWidth" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "InnerDiameterOrWidth"
                    Acustim.Описание_свойства = "Внутренний диаметр"
                    Acustim.Значение = StructObj.InnerDiameterOrWidth
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "InnerLength" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "InnerLength"
                    Acustim.Описание_свойства = "Внутренная длина"
                    Acustim.Значение = StructObj.InnerLength
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '35
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "VerticalPipeClearance" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "VerticalPipeClearance"
                    Acustim.Описание_свойства = "Зазор до верха трубы"
                    Acustim.Значение = StructObj.VerticalPipeClearance
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'смотрим пользовательские свойства
            Dim oData As PartDataField() = StructObj.PartData.GetAllDataFields()
            For Each idp As PartDataField In oData
                boolField = False
                Try
                    Dim Name As String = idp.Name
                    If IsArray(ArrayField) = True Then
                        For i As Integer = 0 To ArrayField.Length - 1
                            If ArrayField(i) Like Name Then
                                boolField = True
                                Exit For
                            End If
                        Next
                    Else
                        boolField = True
                    End If
                    If boolField = True Then
                        Dim desk As String = idp.Description
                        Dim val As String = idp.Value
                        Dim Acustim As Customer = New Customer()
                        Acustim.В_отчет = False
                        Acustim.Имя_свойства = Name
                        Acustim.Описание_свойства = desk
                        Acustim.Значение = val
                        'Acustim.Status = OrderStatus.New
                        custdata.Add(Acustim)
                    End If
                Catch ex As InvalidOperationException
                End Try
            Next
            '==============================================================================================================================
            'выбор трубы
            '==============================================================================================================================
        ElseIf TypeOf obj Is Pipe Then
            custdata.Clear()
            Dim PipeObj As Pipe = obj 'делаем ссылку на трубу
            Dim netWorkStructure As String = PipeObj.NetworkName 'имя трубопроводной сети
            textBox.Text = netWorkStructure
            textBox.Tag = PipeObj.Handle.ToString
            ' 0
            '=========================================================================
            'меняем название кнопки
            Dim TextBlock1 As TextBlock = New TextBlock
            TextBlock1.Text = "Труба"
            TextBlock1.FontSize = 16
            TextBlock1.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            TextBlock1.VerticalAlignment = System.Windows.VerticalAlignment.Center
            button1.Content = TextBlock1
            '===========================================================================
            Dim boolField As Boolean = False
            'имя колодца 0
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Name" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Name"
                    Acustim.Описание_свойства = "Имя элемента"
                    Acustim.Значение = PipeObj.Name
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '1
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Desk" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Desk"
                    Acustim.Описание_свойства = "Описание элемента"
                    Acustim.Значение = PipeObj.Description
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '2
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "RefAlignmentName" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "RefAlignmentName"
                    Acustim.Описание_свойства = "Имя трассы"
                    Acustim.Значение = PipeObj.RefAlignmentName
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '3
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "StartStation" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "StartStation"
                    Acustim.Описание_свойства = "Пикет начала трубы"
                    Acustim.Значение = PipeObj.StartStation
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '3
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "StartOfset" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "StartOfset"
                    Acustim.Описание_свойства = "Смещение начала трубы"
                    Acustim.Значение = PipeObj.StartOffset
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '3
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EndStation" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EndStation"
                    Acustim.Описание_свойства = "Пикет конца трубы"
                    Acustim.Значение = PipeObj.EndStation
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '3
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EndOfset" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EndOfset"
                    Acustim.Описание_свойства = "Смещение конца трубы"
                    Acustim.Значение = PipeObj.EndOffset
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '4
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "StartNorthing" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "StartNorthing"
                    Acustim.Описание_свойства = "Координата X начала трубы"
                    Acustim.Значение = Math.Round(PipeObj.StartPoint.Y, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '5
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "StartEasting" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "StartEasting"
                    Acustim.Описание_свойства = "Координата Y начала трубы"
                    Acustim.Значение = Math.Round(PipeObj.StartPoint.X, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '4
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EndNorthing" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EndNorthing"
                    Acustim.Описание_свойства = "Координата X конца трубы"
                    Acustim.Значение = Math.Round(PipeObj.EndPoint.Y, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '5
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EndEasting" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EndEasting"
                    Acustim.Описание_свойства = "Координата Y конца трубы"
                    Acustim.Значение = Math.Round(PipeObj.EndPoint.X, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '6
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "StartElev" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "StartElev"
                    Acustim.Описание_свойства = "Отметка начала трубы"
                    Acustim.Значение = Math.Round(PipeObj.StartPoint.Z, 3)
                    ' Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '7
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EndElev" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EndElev"
                    Acustim.Описание_свойства = "Отметка конца трубы"
                    Acustim.Значение = Math.Round(PipeObj.EndPoint.Z, 3)
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '8
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "PartDesk" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartDesk"
                    Acustim.Описание_свойства = "Название элемента"
                    Acustim.Значение = PipeObj.PartDescription 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '9
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "PartSizeName" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartSizeName"
                    Acustim.Описание_свойства = "Название размера элемента"
                    Acustim.Значение = PipeObj.PartSizeName 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '10
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "PartSubType" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "PartSubType"
                    Acustim.Описание_свойства = "Описание подтипа"
                    Acustim.Значение = PipeObj.PartSubType 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '10
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "WallThickness" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "WallThickness"
                    Acustim.Описание_свойства = "Толщина стенки"
                    Acustim.Значение = PipeObj.WallThickness 'только чтение
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Bearing" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Bearing"
                    Acustim.Описание_свойства = "Румб"
                    Acustim.Значение = PipeObj.Bearing
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "CrossSectionalShape" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "CrossSectionalShape"
                    Acustim.Описание_свойства = "Форма трубы"
                    Acustim.Значение = PipeObj.CrossSectionalShape.ToString
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "StartStructure" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "StartStructure"
                    Acustim.Описание_свойства = "Начальный колодец"
                    Dim startStructID As ObjectId = PipeObj.StartStructureId
                    If startStructID.IsNull = False Then
                        Dim StartStruct As [Structure] = tr.GetObject(startStructID, OpenMode.ForRead)
                        Acustim.Значение = StartStruct.Name
                        'Acustim.Status = OrderStatus.New
                        custdata.Add(Acustim)
                    End If

                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EndStructure" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EndStructure"
                    Acustim.Описание_свойства = "Конечный колодец"
                    Dim startStructID As ObjectId = PipeObj.EndStructureId
                    If startStructID.IsNull = False Then
                        Dim StartStruct As [Structure] = tr.GetObject(startStructID, OpenMode.ForRead)
                        Acustim.Значение = StartStruct.Name
                        'Acustim.Status = OrderStatus.New
                        custdata.Add(Acustim)
                    End If
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "CoverOfStartPoint" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "CoverOfStartPoint"
                    Acustim.Описание_свойства = "Заглубление начала трубы"
                    Acustim.Значение = PipeObj.CoverOfStartPoint
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "CoverOfEndpoint" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "CoverOfEndpoint"
                    Acustim.Описание_свойства = "Заглубление конца трубы"
                    Acustim.Значение = PipeObj.CoverOfEndpoint
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "MinimumCover" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "MinimumCover"
                    Acustim.Описание_свойства = "Минимальное заглубление"
                    Acustim.Значение = PipeObj.MinimumCover
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "MaximumCover" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "MaximumCover"
                    Acustim.Описание_свойства = "Максимальное заглубление"
                    Acustim.Значение = PipeObj.MaximumCover
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "EnergyGradeLineUp" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "EnergyGradeLineUp"
                    Acustim.Описание_свойства = "Верх энергетического уровня"
                    Acustim.Значение = PipeObj.EnergyGradeLineUp
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "FlowDirection" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "FlowDirection"
                    Acustim.Описание_свойства = "Направление потока"
                    Acustim.Значение = PipeObj.FlowDirection
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "FlowDirectionMethod" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "FlowDirectionMethod"
                    Acustim.Описание_свойства = "Метод определения потока"
                    Acustim.Значение = PipeObj.FlowDirectionMethod.ToString
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "FlowRate" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "FlowRate"
                    Acustim.Описание_свойства = "Скорость потока"
                    Acustim.Значение = PipeObj.FlowRate
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "HoldOnResizeType" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "HoldOnResizeType"
                    Acustim.Описание_свойства = "Изменение размера трубы"
                    Acustim.Значение = PipeObj.HoldOnResizeType.ToString
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "HydraulicGradeLineUp" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "HydraulicGradeLineUp"
                    Acustim.Описание_свойства = "Верх гидравлического уровня"
                    Acustim.Значение = PipeObj.HydraulicGradeLineUp
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "HydraulicGradeLineDown" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "HydraulicGradeLineDown"
                    Acustim.Описание_свойства = "Низ гидравлического уровня"
                    Acustim.Значение = PipeObj.HydraulicGradeLineDown
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "InnerDiameterOrWidth" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "InnerDiameterOrWidth"
                    Acustim.Описание_свойства = "Внутренний диаметр"
                    Acustim.Значение = PipeObj.InnerDiameterOrWidth
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "InnerHeight" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "InnerHeight"
                    Acustim.Описание_свойства = "Внутренняя высота"
                    Acustim.Значение = PipeObj.InnerHeight
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "JunctionLoss" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "JunctionLoss"
                    Acustim.Описание_свойства = "Потеря жидкости в стыке"
                    Acustim.Значение = PipeObj.JunctionLoss
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Length2DCenterToCenter" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Length2DCenterToCenter"
                    Acustim.Описание_свойства = "Длина трубы 2D"
                    Acustim.Значение = PipeObj.Length2DCenterToCenter
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Length3DCenterToCenter" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Length3DCenterToCenter"
                    Acustim.Описание_свойства = "Длина трубы 3D"
                    Acustim.Значение = PipeObj.Length3DCenterToCenter
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Material" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Material"
                    Acustim.Описание_свойства = "Материал трубы"
                    Acustim.Значение = PipeObj.Material
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "OuterDiameterOrWidth" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "OuterDiameterOrWidth"
                    Acustim.Описание_свойства = "Внешний диаметр"
                    Acustim.Значение = PipeObj.OuterDiameterOrWidth
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "OuterHeight" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "OuterHeight"
                    Acustim.Описание_свойства = "Внешняя высота"
                    Acustim.Значение = PipeObj.OuterHeight
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Radius" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Radius"
                    Acustim.Описание_свойства = "Радиус"
                    Acustim.Значение = PipeObj.Radius
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "ReturnPeriod" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "ReturnPeriod"
                    Acustim.Описание_свойства = "Период возврата"
                    Acustim.Значение = PipeObj.ReturnPeriod
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '34
            If IsArray(ArrayField) = True Then
                For i As Integer = 0 To ArrayField.Length - 1
                    If ArrayField(i) Like "Slope" Then
                        boolField = True
                        Exit For
                    End If
                Next
            Else
                boolField = True
            End If
            If boolField = True Then
                Try
                    Dim Acustim As Customer = New Customer()
                    Acustim.В_отчет = False
                    Acustim.Имя_свойства = "Slope"
                    Acustim.Описание_свойства = "Уклон"
                    Acustim.Значение = PipeObj.Slope
                    'Acustim.Status = OrderStatus.New
                    custdata.Add(Acustim)
                Catch ex As InvalidOperationException
                End Try
            End If
            boolField = False
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'смотрим пользовательские свойства
            Dim oData As PartDataField() = PipeObj.PartData.GetAllDataFields()
            For Each idp As PartDataField In oData
                Try
                    boolField = False
                    Dim Name As String = idp.Name
                    If IsArray(ArrayField) = True Then
                        For i As Integer = 0 To ArrayField.Length - 1
                            If ArrayField(i) Like Name Then
                                boolField = True
                                Exit For
                            End If
                        Next
                    Else
                        boolField = True
                    End If
                    If boolField = True Then
                        Dim desk As String = idp.Description
                        Dim val As String = idp.Value
                        Dim Acustim As Customer = New Customer()
                        Acustim.В_отчет = False
                        Acustim.Имя_свойства = Name
                        Acustim.Описание_свойства = desk
                        Acustim.Значение = val
                        'Acustim.Status = OrderStatus.New
                        custdata.Add(Acustim)
                    End If
                Catch ex As InvalidOperationException
                End Try
            Next
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вставляем данные
        dataGrid.ItemsSource = custdata
    End Function
    Public Function FuncSelectionPipe() As Autodesk.AutoCAD.DatabaseServices.Entity
        FuncSelectionPipe = Nothing
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim acSSPrompt As PromptSelectionResult = m_Editor.SelectImplied()
        Dim obj1 As Autodesk.AutoCAD.DatabaseServices.Entity = Nothing
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                For Each id As ObjectId In acSSet.GetObjectIds
                    Dim objTemp As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                    If TypeOf objTemp Is [Structure] Or TypeOf objTemp Is Pipe Then
                        obj1 = objTemp
                    End If
                Next
            End If
            If IsNothing(obj1) = False Then
                FuncSelectionPipe = obj1
                Dim StrFilter As String = ComboBoxFilter.Text
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'читаем фильтр
                Dim ArrayField As String() = Nothing
                Dim CountArrayField As Integer = 0
                Dim NameFilter As String = ComboBoxFilter.Text
                Dim NAmefileF As String = ComboBoxFilter.Tag & NameFilter & ".ini"
                'читаем поля
                If IO.File.Exists(ComboBoxFilter.Tag & NameFilter & ".ini") = True Then
                    Using sr As IO.StreamReader = New IO.StreamReader(NAmefileF)
                        Dim line As String
                        Do
                            line = sr.ReadLine
                            If line Is Nothing Then
                                Exit Do
                            End If
                            If Not (line Like "") Then
                                ReDim Preserve ArrayField(CountArrayField)
                                ArrayField(CountArrayField) = line
                                CountArrayField += 1
                            End If
                        Loop
                    End Using
                End If
                Dim boolPipe As Boolean = FuncReadPropertiesPipe(obj1, tr, ArrayField)
            End If
        End Using
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменение размера палитры
    Private Sub PipeProperties_SizeChanged(sender As Object, e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        Try
            Dim s1 As Double = e.NewSize.Height 'ширина формы (палитры)
            Dim s2 As Double = e.NewSize.Width
            groupBox.Height = s1 - 80
            groupBox.Width = s2 - 5
            dataGrid.Height = s1 - 150
            dataGrid.Width = s2 - 20
        Catch ex As Exception
        End Try
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ввод данных (Одиночное нажатие переключателя
    Private Sub dataGrid_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles dataGrid.SelectedCellsChanged
        Dim columnSelect As DataGridColumn = dataGrid.CurrentCell.Column 'номер выбранного столбца
        Dim cellSelect As IList(Of DataGridCellInfo) = dataGrid.SelectedCells
        Dim ind As Integer = 1
        If columnSelect.DisplayIndex = 1 Then
            For Each cellS As DataGridCellInfo In cellSelect
                Dim bool As PipeLinePalette.Customer = cellS.Item
                If ind = 5 Then
                    If bool.В_отчет = True Then
                        bool.В_отчет = False
                        If bool.Поз = indCount Then
                            indCount -= 1
                            bool.Поз = 0
                        Else
                            bool.Поз = 0
                        End If
                    Else
                        bool.В_отчет = True
                        indCount += 1
                        bool.Поз = indCount
                    End If
                    ind = 1
                Else
                    ind += 1
                End If
            Next
            Dim custdata2 As ObservableCollection(Of Customer) = New ObservableCollection(Of Customer)
            Dim Acustim2 As Customer = New Customer()
            Acustim2.В_отчет = False
            Acustim2.Имя_свойства = ""
            Acustim2.Описание_свойства = ""
            Acustim2.Значение = ""
            custdata2.Add(Acustim2)
            dataGrid.ItemsSource = custdata2
            dataGrid.ItemsSource = custdata
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'множественный выбор
    Private Sub dataGrid_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles dataGrid.MouseDown
        Dim cellSelect As IList(Of DataGridCellInfo) = dataGrid.SelectedCells
        Dim indSelect As Object = dataGrid.SelectedIndex
        Dim ind As Integer = 1
        If cellSelect.Count > 6 Then
            For Each cellS As DataGridCellInfo In cellSelect
                Dim bool As PipeLinePalette.Customer = cellS.Item
                If ind = 5 Then
                    If bool.В_отчет = True Then
                        bool.В_отчет = False
                    Else
                        bool.В_отчет = True
                        indCount += 1
                        bool.Поз = indCount
                    End If
                    ind = 1
                Else
                    ind += 1
                End If
            Next
        End If
        Dim custdata2 As ObservableCollection(Of Customer) = New ObservableCollection(Of Customer)
        Dim Acustim2 As Customer = New Customer()
        Acustim2.В_отчет = False
        Acustim2.Имя_свойства = ""
        Acustim2.Описание_свойства = ""
        Acustim2.Значение = ""
        custdata2.Add(Acustim2)
        dataGrid.ItemsSource = custdata2
        dataGrid.ItemsSource = custdata
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор объекта
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Private Sub button_Click(sender As Object, e As RoutedEventArgs) Handles button1.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        AddHandler acDoc.ImpliedSelectionChanged, AddressOf FuncSelectionPipe
        Using tr
            Dim ArrayStructExcel As String(,) = Nothing
            Dim CountArrayStructExcel As Integer = 0
            Dim resBlk As PromptEntityResult = Nothing
            Dim opt As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите объект: ")
            resBlk = m_Editor.GetEntity(opt)
            If resBlk.Status = PromptStatus.Cancel Or resBlk.Status = PromptStatus.Error Or resBlk.Status = PromptStatus.None Then
                Exit Sub
            End If
            Dim pvId As ObjectId = resBlk.ObjectId
            Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pvId, OpenMode.ForRead)


            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'читаем фильтр
            Dim ArrayField As String() = Nothing
            Dim CountArrayField As Integer = 0
            Dim NameFilter As String = ComboBoxFilter.Text
            Dim NAmefileF As String = ComboBoxFilter.Tag & NameFilter & ".ini"
            'читаем поля
            If IO.File.Exists(ComboBoxFilter.Tag & NameFilter & ".ini") = True Then
                Using sr As IO.StreamReader = New IO.StreamReader(NAmefileF)
                    Dim line As String
                    Do
                        line = sr.ReadLine
                        If line Is Nothing Then
                            Exit Do
                        End If
                        If Not (line Like "") Then
                            ReDim Preserve ArrayField(CountArrayField)
                            ArrayField(CountArrayField) = line
                            CountArrayField += 1
                        End If
                    Loop
                End Using
            End If
            Dim boolPipeProperties As Boolean = FuncReadPropertiesPipe(obj, tr, ArrayField)
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'сформировать отчет
    Private Sub button1_Click(sender As Object, e As RoutedEventArgs) Handles button1.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        'смотрим точность
        Dim SistPeremPK As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value 'настройки пикета
        Dim SistPeremDist As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Coordinate.Precision.Value 'точность по координатам
        Dim SistPeremRm As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Distance.Precision.Value 'точность по координатам
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Dim docLock As DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            'делаем ссылку на поля которые выбраны
            Dim TempArrayName As String(,) = Nothing
            Dim ArrayName As String(,) = Nothing
            Dim CountArrayName As Integer = 0
            If dataGrid.Items.Count > 0 Then
                For i As Integer = 0 To dataGrid.Items.Count - 2
                    Dim cust As Customer = dataGrid.Items.Item(i)
                    If cust.В_отчет = True Then
                        ReDim Preserve TempArrayName(3, CountArrayName)
                        TempArrayName(0, CountArrayName) = cust.Поз
                        TempArrayName(1, CountArrayName) = cust.Имя_свойства
                        TempArrayName(2, CountArrayName) = cust.Описание_свойства
                        TempArrayName(3, CountArrayName) = cust.Значение
                        CountArrayName += 1
                    End If
                Next
                If IsArray(TempArrayName) = True Then
                    'сортируем массив
                    Dim boolSort As Boolean = MathFunction.FuncSortStrArray(TempArrayName, 0)
                    For j As Integer = 0 To TempArrayName.GetUpperBound(1)
                        ReDim Preserve ArrayName(3, j)
                        ArrayName(0, j) = TempArrayName(1, j)
                        ArrayName(1, j) = TempArrayName(2, j)
                        ArrayName(2, j) = TempArrayName(3, j)
                    Next
                End If
            End If
            If IsArray(ArrayName) = True Then
                If comboBox.Text Like "Создать фильтр" Then
                    Dim nameFileFil As String = InputBox("Имя фильтра", "Создать фильтр")
                    Dim FullNameFile As String = ComboBoxFilter.Tag & nameFileFil & ".ini"
                    Using sr As IO.StreamWriter = New IO.StreamWriter(FullNameFile, IO.FileMode.Create, System.Text.Encoding.UTF8)
                        For j As Integer = 0 To TempArrayName.GetUpperBound(1)
                            sr.WriteLine(ArrayName(0, j))
                        Next
                        sr.Close()
                        sr.Dispose()
                    End Using
                    Exit Sub
                End If
                Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'запрос на обработку всей выбранной полилинии или только ее части
                Dim pIntOpts As PromptKeywordOptions = New PromptKeywordOptions("")
                pIntOpts.Message = vbCrLf & "Выбор объектов трубопроводных сетей для отчета "
                pIntOpts.Keywords.Add("Все")
                pIntOpts.Keywords.Add("Сеть")
                pIntOpts.Keywords.Add("Объект")
                pIntOpts.Keywords.Add("Нет")
                pIntOpts.Keywords.Default = "Сеть"
                pIntOpts.AllowNone = True
                Dim pIntRes As PromptResult = acDoc.Editor.GetKeywords(pIntOpts)
                If pIntRes.Status = PromptStatus.Cancel Or pIntRes.Status = PromptStatus.Error Then
                    Throw New System.Exception("Прервано пользователем")
                    Exit Sub 'если нажали ESC то выходим из программы
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'заводим коллекцию по колодцам
                Dim idCollectionStructure As ObjectIdCollection = New ObjectIdCollection
                If pIntRes.StringResult Like "Все" Then 'все колодцы чертежа
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'забираем все трубопроводные сети
                    Dim netPipes As ObjectIdCollection = CivilDoc.GetPipeNetworkIds
                    If netPipes.Count > 0 Then
                        For Each IDnetPipe As ObjectId In netPipes
                            Dim netPipe As Network = tr.GetObject(IDnetPipe, OpenMode.ForRead)
                            Dim TextBlock1 As TextBlock = button1.Content
                            If TextBlock1.Text Like "Колодец" Then
                                Dim TempIdStruct As ObjectIdCollection = netPipe.GetStructureIds
                                For Each id As ObjectId In TempIdStruct
                                    idCollectionStructure.Add(id)
                                Next
                            Else
                                Dim TempIdPipe As ObjectIdCollection = netPipe.GetPipeIds
                                For Each id As ObjectId In TempIdPipe
                                    idCollectionStructure.Add(id)
                                Next
                            End If

                        Next
                    End If
                ElseIf pIntRes.StringResult Like "Сеть" Then 'определенная сеть
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'забираем все трубопроводные сети
                    Dim netPipes As ObjectIdCollection = CivilDoc.GetPipeNetworkIds
                    If netPipes.Count > 0 Then
                        For Each IDnetPipe As ObjectId In netPipes
                            Dim netPipe As Network = tr.GetObject(IDnetPipe, OpenMode.ForRead)
                            Dim TextBlock1 As TextBlock = button1.Content
                            If TextBlock1.Text Like "Колодец" Then
                                If netPipe.Name Like textBox.Text Then
                                    idCollectionStructure = netPipe.GetStructureIds
                                    Exit For
                                End If
                            Else
                                If netPipe.Name Like textBox.Text Then
                                    idCollectionStructure = netPipe.GetPipeIds
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                ElseIf pIntRes.StringResult Like "Объект" Then 'отдельный колодец
                    Dim TextBlock1 As TextBlock = button1.Content
                    Dim acSSPrompt As PromptSelectionResult = m_Editor.GetSelection()
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim acSSet As SelectionSet = acSSPrompt.Value
                        For Each id As ObjectId In acSSet.GetObjectIds
                            Dim ent As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(id, OpenMode.ForRead)
                            If TextBlock1.Text Like "Колодец" And TypeOf ent Is [Structure] Then
                                idCollectionStructure.Add(id)
                            ElseIf TextBlock1.Text Like "Труба" And TypeOf ent Is Pipe Then
                                idCollectionStructure.Add(id)
                            End If
                        Next
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'формируем массив для отчета
                Dim un As Autodesk.AutoCAD.Runtime.DistanceUnitFormat = Autodesk.AutoCAD.Runtime.DistanceUnitFormat.Decimal
                Dim ArrayExcel As String(,) = Nothing
                Dim CountArrayExcel As Integer = 0
                If idCollectionStructure.Count > 0 Then
                    Dim TextBlock1 As TextBlock = button1.Content
                    If TextBlock1.Text Like "Колодец" Then
                        For Each idStruct As ObjectId In idCollectionStructure
                            Dim StructObj As [Structure] = tr.GetObject(idStruct, OpenMode.ForRead)
                            ReDim Preserve ArrayExcel(CountArrayName, CountArrayExcel)
                            ArrayExcel(CountArrayName, CountArrayExcel) = StructObj.Handle.ToString
                            For i As Integer = 0 To ArrayName.GetUpperBound(1)
                                Dim Name As String = ArrayName(0, i)
                                'имя колодца 0
                                Try
                                    If Name Like "Name" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.Name
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '1
                                Try
                                    If Name Like "Desk" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.Description
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '2
                                Try
                                    If Name Like "Northing" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Northing, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '3
                                Try
                                    If Name Like "Easting" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Easting, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '4
                                Try
                                    If Name Like "ElevTop" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.RimElevation, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '5
                                Try
                                    If Name Like "ElevDown" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.SumpElevation, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '6
                                Try
                                    If Name Like "PartDesk" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.PartDescription 'только чтение
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '7
                                Try
                                    If Name Like "PartSizeNAme" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.PartSizeName 'только чтение
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '8
                                Try
                                    If Name Like "PartSubType" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.PartSubType 'только чтение
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '9
                                Try
                                    If Name Like "Diameter" Then
                                        Dim d As Double = StructObj.DiameterOrWidth 'только чтение
                                        If d < 20 Then
                                            d = d * 1000
                                            ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(d, un, 0)
                                        Else
                                            ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(d, un, 3)
                                        End If
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '10
                                Try
                                    If Name Like "Lenght" Then
                                        Dim fc As Double = StructObj.Length
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '11
                                Try
                                    If Name Like "RimToSumpHeight" Then
                                        Dim fc As Double = StructObj.RimToSumpHeight
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '12
                                Try
                                    If Name Like "Height" Then
                                        Dim fc As Double = StructObj.Height
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '13
                                Try
                                    If Name Like "BarrelPipeClearance" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.BarrelPipeClearance, un, SistPeremRm)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '14
                                Try
                                    If Name Like "SurfElevAtInsPoint" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.SurfaceElevationAtInsertionPoint, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As CivilException
                                End Try
                                '15
                                Try
                                    If Name Like "ConeHeight" Then
                                        Dim fc As Double = StructObj.ConeHeight
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '16
                                Try
                                    If Name Like "ConnectedPartCount" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.ConnectedPartCount
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '17
                                Try
                                    If Name Like "SurfaceAdjustmentValue" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.SurfaceAdjustmentValue, un, SistPeremRm)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '18
                                Try
                                    If Name Like "ConnectedPipesCount" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.ConnectedPipesCount
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '19
                                'Try
                                '    If Name Like "Domain" Then
                                '        ArrayExcel(i, CountArrayExcel) = StructObj.Domain
                                '    End If
                                'Catch ex As InvalidOperationException
                                'End Try
                                '20
                                Try
                                    If Name Like "FloorThickness" Then
                                        Dim fc As Double = StructObj.FloorThickness
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '21
                                Try
                                    If Name Like "RuleSetStyleName" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.RuleSetStyleName 'стиль колодца
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '22
                                Try
                                    If Name Like "FrameDiameter" Then
                                        Dim fc As Double = StructObj.FrameDiameter
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If

                                Catch ex As InvalidOperationException
                                End Try
                                '23
                                Try
                                    If Name Like "FrameHeight" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.FrameHeight, un, SistPeremRm)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '24
                                Try
                                    If Name Like "Grate" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.Grate 'тип решотки
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '25
                                Try
                                    If Name Like "HeadwallBaseThickness" Then
                                        Dim fc As Double = StructObj.HeadwallBaseThickness
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '26
                                Try
                                    If Name Like "RefAlignmentName" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.RefAlignmentName
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '27
                                Try
                                    If Name Like "HeadwallBaseWidth" Then
                                        Dim fc As Double = StructObj.HeadwallBaseWidth
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '28
                                Try
                                    If Name Like "InnerDiameterOrWidth" Then
                                        Dim fc As Double = StructObj.InnerDiameterOrWidth
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '29
                                Try
                                    If Name Like "InnerLength" Then
                                        Dim fc As Double = StructObj.InnerLength
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '30
                                Try
                                    If Name Like "NetworkName" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.NetworkName
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '31
                                Try
                                    If Name Like "Offset" Then
                                        Dim ofsetPk As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Offset, un, SistPeremPK)
                                        If ofsetPk < 0 Then
                                            ArrayExcel(i, CountArrayExcel) = "Лево " & Math.Abs(ofsetPk).ToString
                                        ElseIf ofsetPk > 0 Then
                                            ArrayExcel(i, CountArrayExcel) = "Право " & Math.Abs(ofsetPk).ToString
                                        Else
                                            ArrayExcel(i, CountArrayExcel) = 0
                                        End If
                                        Continue For
                                    End If
                                Catch ex As CivilException
                                End Try
                                '32
                                Try
                                    If Name Like "VerticalPipeClearance" Then
                                        Dim fc As Double = StructObj.VerticalPipeClearance
                                        fc = fc * 1000
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(fc, un, 0)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '33
                                Try
                                    If Name Like "RimElevation" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.RimElevation, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Rotation" Then
                                        Dim an As Autodesk.AutoCAD.Runtime.AngularUnitFormat = Autodesk.AutoCAD.Runtime.AngularUnitFormat.DegreesMinutesSeconds
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.AngleToString(StructObj.Rotation, an, 6)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '35
                                Try
                                    If Name Like "Station" Then
                                        Dim PK As Double = StructObj.Station
                                        Dim intPK As Integer = Int(PK / 100)
                                        Dim DblPk As Double = PK - intPK * 100
                                        Dim DblPKFormat As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(DblPk, un, SistPeremPK)
                                        ArrayExcel(i, CountArrayExcel) = intPK & "+" & DblPKFormat
                                        Continue For
                                    End If
                                Catch ex As CivilException
                                End Try
                                '36
                                Try
                                    If Name Like "SumpDepth" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.SumpDepth, un, SistPeremDist)
                                        Continue For
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'смотрим пользовательские свойства
                                Dim oData As PartDataField() = StructObj.PartData.GetAllDataFields()
                                For Each idp As PartDataField In oData
                                    Try
                                        Dim NameTemp As String = idp.Name
                                        Dim val As String = idp.Value
                                        If Name Like NameTemp Then
                                            Dim un1 As String = idp.Units
                                            If un1 Like "m" Then
                                                ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(val, un, SistPeremDist)
                                            ElseIf un1 Like "mm" Then
                                                ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(val, un, 0)
                                            Else
                                                ArrayExcel(i, CountArrayExcel) = val
                                            End If
                                            Continue For
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                Next
                            Next i
                            CountArrayExcel += 1
                        Next
                        '=====================================================================================================
                        'характеристики по трубам
                        '=====================================================================================================
                    Else
                        For Each idStruct As ObjectId In idCollectionStructure
                            Dim StructObj As Pipe = tr.GetObject(idStruct, OpenMode.ForRead)
                            ReDim Preserve ArrayExcel(CountArrayName, CountArrayExcel)
                            ArrayExcel(CountArrayName, CountArrayExcel) = StructObj.Handle.ToString
                            For i As Integer = 0 To ArrayName.GetUpperBound(1)
                                Dim Name As String = ArrayName(0, i)
                                '===========================================================================
                                Try
                                    If Name Like "Name" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.Name
                                    End If
                                Catch ex As CivilException
                                End Try
                                '1
                                Try
                                    If Name Like "Desk" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.Description
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '2
                                Try
                                    If Name Like "RefAlignmentName" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.RefAlignmentName
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '3
                                Try
                                    If Name Like "StartStation" Then
                                        Dim PK As Double = StructObj.StartStation
                                        Dim intPK As Integer = Int(PK / 100)
                                        Dim DblPk As Double = PK - intPK * 100
                                        Dim DblPKFormat As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(DblPk, un, SistPeremPK)
                                        ArrayExcel(i, CountArrayExcel) = intPK & "+" & DblPKFormat
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '3
                                Try
                                    If Name Like "StartOfset" Then
                                        Dim ofsetPk As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.StartOffset, un, SistPeremPK)
                                        If ofsetPk < 0 Then
                                            ArrayExcel(i, CountArrayExcel) = "Лево " & Math.Abs(ofsetPk).ToString
                                        ElseIf ofsetPk > 0 Then
                                            ArrayExcel(i, CountArrayExcel) = "Право " & Math.Abs(ofsetPk).ToString
                                        Else
                                            ArrayExcel(i, CountArrayExcel) = 0
                                        End If
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '3
                                Try
                                    If Name Like "EndStation" Then
                                        Dim PK As Double = StructObj.EndStation
                                        Dim intPK As Integer = Int(PK / 100)
                                        Dim DblPk As Double = PK - intPK * 100
                                        Dim DblPKFormat As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(DblPk, un, SistPeremPK)
                                        ArrayExcel(i, CountArrayExcel) = intPK & "+" & DblPKFormat
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '3
                                Try
                                    If Name Like "EndOfset" Then
                                        Dim ofsetPk As Double = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.EndOffset, un, SistPeremPK)
                                        If ofsetPk < 0 Then
                                            ArrayExcel(i, CountArrayExcel) = "Лево " & Math.Abs(ofsetPk).ToString
                                        ElseIf ofsetPk > 0 Then
                                            ArrayExcel(i, CountArrayExcel) = "Право " & Math.Abs(ofsetPk).ToString
                                        Else
                                            ArrayExcel(i, CountArrayExcel) = 0
                                        End If
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '4
                                Try
                                    If Name Like "StartNorthing" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.StartPoint.Y, un, SistPeremDist)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '5
                                Try
                                    If Name Like "StartEasting" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.StartPoint.X, un, SistPeremDist)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '4
                                Try
                                    If Name Like "EndNorthing" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.EndPoint.Y, un, SistPeremDist)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '5
                                Try
                                    If Name Like "EndEasting" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.EndPoint.X, un, SistPeremDist)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '6
                                Try
                                    If Name Like "StartElev" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.StartPoint.Z, un, SistPeremDist)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '7
                                Try
                                    If Name Like "EndElev" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.EndPoint.Z, un, SistPeremDist)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '8
                                Try
                                    If Name Like "PartDesk" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.PartDescription
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '9
                                Try
                                    If Name Like "PartSizeNAme" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.PartSizeName
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '10
                                Try
                                    If Name Like "PartSubType" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.PartSubType
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '10
                                Try
                                    If Name Like "WallThickness" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.WallThickness
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Bearing" Then
                                        Dim rumb As Double = StructObj.Bearing

                                        'ArrayExcel(i, CountArrayExcel) = 

                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "CrossSectionalShape" Then
                                        Dim sv As SweptShapeType = StructObj.CrossSectionalShape
                                        If sv = SweptShapeType.Arched Then
                                            ArrayExcel(i, CountArrayExcel) = "Арочная"
                                        ElseIf sv = SweptShapeType.Circular Then
                                            ArrayExcel(i, CountArrayExcel) = "Круглая"
                                        ElseIf sv = SweptShapeType.CustomShape Then
                                            ArrayExcel(i, CountArrayExcel) = "Пользовательская"
                                        ElseIf sv = SweptShapeType.EggShaped Then
                                            ArrayExcel(i, CountArrayExcel) = "Пользовательская"
                                        ElseIf sv = SweptShapeType.Elliptical Then
                                            ArrayExcel(i, CountArrayExcel) = "Эллипс"
                                        ElseIf sv = SweptShapeType.HorizontalElliptical Then
                                            ArrayExcel(i, CountArrayExcel) = "Эллипс на горизонтали"
                                        ElseIf sv = SweptShapeType.Rectangular Then
                                            ArrayExcel(i, CountArrayExcel) = "Прямоугольная"
                                        ElseIf sv = SweptShapeType.Undefined Then
                                            ArrayExcel(i, CountArrayExcel) = "Неопределена"
                                        End If
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "StartStructure" Then
                                        Dim startStructID As ObjectId = StructObj.StartStructureId
                                        Dim StartStruct As [Structure] = tr.GetObject(startStructID, OpenMode.ForRead)
                                        ArrayExcel(i, CountArrayExcel) = StartStruct.Name
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "EndStructure" Then
                                        Dim startStructID As ObjectId = StructObj.EndStructureId
                                        Dim StartStruct As [Structure] = tr.GetObject(startStructID, OpenMode.ForRead)
                                        ArrayExcel(i, CountArrayExcel) = StartStruct.Name
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "CoverOfStartPoint" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.CoverOfStartPoint, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "MinimumCover" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.MinimumCover, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "MaximumCover" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.MaximumCover, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "CoverOfEndpoint" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.CoverOfEndpoint, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                'Try
                                '    If Name Like "Domain" Then
                                '        ArrayExcel(i, CountArrayExcel) = StructObj.Domain
                                '    End If
                                'Catch ex As InvalidOperationException
                                'End Try
                                '34
                                Try
                                    If Name Like "EnergyGradeLineUp" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.EnergyGradeLineUp, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "FlowDirection" Then
                                        Dim fd As FlowDirectionType = StructObj.FlowDirection
                                        If fd = FlowDirectionType.Bidirectional Then
                                            ArrayExcel(i, CountArrayExcel) = "Двунаправленный"
                                        ElseIf fd = FlowDirectionType.EndToStart Then
                                            ArrayExcel(i, CountArrayExcel) = "От конца к началу трубы"
                                        ElseIf fd = FlowDirectionType.StartToEnd Then
                                            ArrayExcel(i, CountArrayExcel) = "От начала к концу трубы"
                                        End If
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "FlowDirectionMethod" Then
                                        Dim fdm As FlowDirectionMethodType = StructObj.FlowDirectionMethod
                                        If fdm = FlowDirectionMethodType.Bidirectional Then
                                            ArrayExcel(i, CountArrayExcel) = "Двунаправленный"
                                        ElseIf fdm = FlowDirectionMethodType.BySlope Then
                                            ArrayExcel(i, CountArrayExcel) = "Уклон"
                                        ElseIf fdm = FlowDirectionMethodType.EndToStart Then
                                            ArrayExcel(i, CountArrayExcel) = "От конца к началу трубы"
                                        ElseIf fdm = FlowDirectionMethodType.StartToEnd Then
                                            ArrayExcel(i, CountArrayExcel) = "От начала к концу трубы"
                                        End If
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "FlowRate" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.FlowRate, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "HoldOnResizeType" Then
                                        Dim HOnRT As HoldOnResizeType = StructObj.HoldOnResizeType
                                        If HOnRT = HoldOnResizeType.Centerline Then
                                            ArrayExcel(i, CountArrayExcel) = "Центр трубы"
                                        ElseIf HOnRT = HoldOnResizeType.Crown Then
                                            ArrayExcel(i, CountArrayExcel) = "Верх трубы"
                                        ElseIf HOnRT = HoldOnResizeType.Invert Then
                                            ArrayExcel(i, CountArrayExcel) = "Низ трубы"
                                        End If
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "HydraulicGradeLineUp" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.HydraulicGradeLineUp, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "HydraulicGradeLineDown" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.HydraulicGradeLineDown, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "InnerDiameterOrWidth" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.InnerDiameterOrWidth * 1000, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "InnerHeight" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.InnerHeight * 1000, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "JunctionLoss" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.JunctionLoss, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Length2DCenterToCenter" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Length2DCenterToCenter, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Length3DCenterToCenter" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Length3DCenterToCenter, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Material" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.Material
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "OuterDiameterOrWidth" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.OuterDiameterOrWidth * 1000, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "OuterHeight" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.OuterHeight * 1000, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Radius" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Radius * 1000, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "ReturnPeriod" Then
                                        ArrayExcel(i, CountArrayExcel) = StructObj.ReturnPeriod
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '34
                                Try
                                    If Name Like "Slope" Then
                                        ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(StructObj.Slope * 1000, un, SistPeremRm)
                                    End If
                                Catch ex As InvalidOperationException
                                End Try
                                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                'смотрим пользовательские свойства
                                Dim oData As PartDataField() = StructObj.PartData.GetAllDataFields()
                                For Each idp As PartDataField In oData
                                    Try
                                        Dim NameTemp As String = idp.Name
                                        Dim val As String = idp.Value
                                        If Name Like NameTemp Then
                                            Dim un1 As String = idp.Units
                                            If un1 Like "m" Then
                                                ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(val, un, SistPeremDist)
                                            ElseIf un1 Like "mm" Then
                                                ArrayExcel(i, CountArrayExcel) = Autodesk.AutoCAD.Runtime.Converter.DistanceToString(val, un, 0)
                                            Else
                                                ArrayExcel(i, CountArrayExcel) = val
                                            End If
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                Next
                            Next i
                            CountArrayExcel += 1
                        Next
                    End If
                Else
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'одиночный колодец
                    ReDim Preserve ArrayExcel(CountArrayName, 0)
                    ArrayExcel(CountArrayName, 0) = textBox.Tag
                    For i As Integer = 0 To ArrayName.GetUpperBound(1)
                        If dataGrid.Items.Count > 0 Then
                            For j As Integer = 0 To dataGrid.Items.Count - 1
                                Dim cust As Customer = dataGrid.Items.Item(i)
                                If cust.Имя_свойства Like ArrayName(0, i) Then
                                    ArrayExcel(i, CountArrayExcel) = cust.Значение
                                End If
                            Next
                        End If
                    Next
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'ВЫВОДИМ ОТЧЕТ
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                If IsArray(ArrayExcel) = True Then
                    If comboBox.Text Like "Таблица Excel" Then
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'запускаем Excel и формируем шапку таблицы
                        'находим все директории с файлами
                        'запускаем эксель
                        Dim xlApp As Excel.Application = Nothing
                        Dim xlBook As Excel.Workbook = Nothing
                        Dim xlSheets As Excel.Worksheet = Nothing
                        Dim cell1 As Excel.Range = Nothing 'первая ячейка в строке
                        Dim cell2 As Excel.Range = Nothing 'последняя ячейка в строке
                        Dim CellRng As Excel.Range = Nothing 'массив ячеек
                        xlApp = CreateObject("Excel.Application")
                        xlBook = xlApp.Workbooks.Add()
                        xlApp.Visible = True
                        'находим нужный лист(Ведомость объемов работ)
                        Dim xlsheet As Excel.Worksheet = xlBook.Sheets.Item("Лист1")
                        xlsheet.Activate()
                        '==================================================================================
                        'формируем шапку таблицы
                        '===============================================================================
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'верхняя ячейка
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'выводим столбец с пикетажем
                        xlsheet.Cells(1, 1) = "'Номер п/п" 'порядковый номер
                        For i As Integer = 0 To ArrayName.GetUpperBound(1)
                            Dim valTbl As String = ArrayName(1, i) 'описание свойства
                            If IsNothing(valTbl) = True Then
                                valTbl = "Значение не определено"
                            End If
                            xlsheet.Cells(1, i + 2) = valTbl
                        Next i
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'заполняем свойства по колодцам
                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                            xlsheet.Cells(i + 2, 1) = i + 1 'порядковый номер
                            For j As Integer = 0 To ArrayName.GetUpperBound(1)
                                Dim valTbl As String = ArrayExcel(j, i) 'описание свойства
                                If IsNothing(valTbl) = True Then
                                    valTbl = "Значение не определено"
                                End If
                                xlsheet.Cells(i + 2, j + 2) = valTbl 'значение
                            Next j
                        Next i
                        xlApp = Nothing
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    ElseIf comboBox.Text Like "Таблица автокад" Then
                        'вставляем слой для тангенциального хода и номеров вершин углов
                        Dim NameLayer As String = "1-Ведомости"
                        Dim IDLayerIns As ObjectId = AcadFunction.FuncAddLayer(NameLayer, 7,, 20)
                        If IsNothing(IDLayerIns) = True Then
                            NameLayer = "0"
                        ElseIf IDLayerIns.IsNull = True Then
                            NameLayer = "0"
                        End If
                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                        pPtOpts.Message = vbLf & "Укажите точку вставки таблицы: "
                        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                        Dim ptStart As Point3d = pPtRes.Value
                        If pPtRes.Status = PromptStatus.Cancel Then GoTo lineend
                        Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = New Autodesk.AutoCAD.DatabaseServices.Table()
                        tb.TableStyle = acCurDb.Tablestyle
                        tb.SetDatabaseDefaults()
                        acBlkTblRec.AppendEntity(tb)
                        tb.Position = ptStart
                        tb.Layer = NameLayer
                        tb.SetSize(3, ArrayName.GetUpperBound(1) + 2)
                        tb.SetRowHeight(8)
                        tb.TextHeight(20)
                        tb.Cells(0, 0).Value = "Экспликация объектов трубопроводной сети"
                        tb.Cells(1, 0).Value = "Номер п\п"
                        For i As Integer = 0 To ArrayName.GetUpperBound(1)
                            Dim valTbl As String = ArrayName(1, i) 'описание свойства
                            If IsNothing(valTbl) = True Then
                                valTbl = "Значение не определено"
                            End If
                            tb.Cells(1, i + 1).Value = valTbl 'описание свойства
                        Next
                        Dim NumRows As Integer = 2
                        For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                            tb.InsertRows(NumRows + i + 1, 2.5, 1)
                            tb.Cells(NumRows + i, 0).Value = i + 1
                            tb.Columns.Item(0).Width = 15
                            For j As Integer = 0 To ArrayName.GetUpperBound(1)
                                Dim valTbl As String = ArrayExcel(j, i) 'описание свойства
                                If IsNothing(valTbl) = True Then
                                    valTbl = "Значение не определено"
                                End If
                                If valTbl = "" Then valTbl = "-"
                                tb.Cells(NumRows + i, j + 1).Value = valTbl
                                tb.Columns.Item(j + 1).Width = 50
                            Next j
                        Next i
                        tr.AddNewlyCreatedDBObject(tb, True)
lineend:
                    ElseIf comboBox.Text Like "Сделать Solid3d" Then
                        'достаем handle
                        '===============================================================================
                        'вставляем слой
                        'вставляем слой для тангенциального хода и номеров вершин углов
                        Dim NameLayer As String = "ИС-Объекты трубопроводной сети"
                        Dim IDLayerIns As ObjectId = AcadFunction.FuncAddLayer(NameLayer, 7,, 20)
                        If IsNothing(IDLayerIns) = True Then
                            NameLayer = "0"
                        ElseIf IDLayerIns.IsNull = True Then
                            NameLayer = "0"
                        End If
                        Dim TextBlock1 As TextBlock = button1.Content
                        If TextBlock1.Text Like "Колодец" Then
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                Dim hg As String = ArrayExcel(CountArrayName, i)
                                Dim IDObj As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                Dim StructObj As [Structure] = tr.GetObject(IDObj, OpenMode.ForRead)
                                Dim Solid3dObj As Solid3d = StructObj.Solid3dBody
                                Dim blkName As String = "Struct_" & StructObj.Name.Trim & "_" & hg
                                Using acBlkTblRecTemp As New BlockTableRecord
                                    If Not acBlkTbl.Has(blkName) Then
                                        acBlkTblRecTemp.Name = blkName
                                        acBlkTblRecTemp.Origin = StructObj.Position
                                        acBlkTblRecTemp.AppendEntity(Solid3dObj)
                                        acBlkTbl.UpgradeOpen()
                                        acBlkTbl.Add(acBlkTblRecTemp)
                                        Dim br As BlockReference = New BlockReference(StructObj.Position, acBlkTblRecTemp.Id)
                                        Dim btr As BlockTableRecord = tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                                        br.Layer = NameLayer
                                        br.ScaleFactors = New Scale3d(1, 1, 1)
                                        br.Rotation = 0
                                        btr.AppendEntity(br)
                                        tr.AddNewlyCreatedDBObject(br, True)
                                        tr.AddNewlyCreatedDBObject(acBlkTblRecTemp, True)
                                        '=======================================================================================
                                        'добавляем атрибуты
                                        For j As Integer = 0 To ArrayName.GetUpperBound(1)
                                            Dim Name As String = ArrayName(0, j)
                                            Dim Desk As String = ArrayName(1, j)
                                            If IsNothing(Desk) = True Then Desk = "Не_определено"
                                            If Desk = "" Then Desk = "Не_определено"
                                            Dim val As String = ArrayExcel(j, i)
                                            If IsNothing(val) = True Then val = "Не определено"
                                            If val = "" Then val = "Не определено"
                                            Desk = Desk.Replace(" ", "_")
                                            Dim ad As AttributeDefinition = New AttributeDefinition
                                            ad.SetDatabaseDefaults()
                                            ' btr.AppendEntity(ad)
                                            Dim ar As AttributeReference = New AttributeReference()
                                            ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                            br.AttributeCollection.AppendAttribute(ar)
                                            tr.AddNewlyCreatedDBObject(ar, True)
                                            ar.Layer = NameLayer
                                            ar.Tag = Desk
                                            ar.TextString = val
                                            ar.Invisible = False
                                            ar.Oblique = 0
                                            ar.Invisible = False
                                        Next j
                                    End If
                                End Using
                            Next i
                        Else
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                Dim hg As String = ArrayExcel(CountArrayName, i)
                                Dim IDObj As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                Dim StructObj As Pipe = tr.GetObject(IDObj, OpenMode.ForRead)
                                Dim Solid3dObj As Solid3d = StructObj.Solid3dBody
                                Dim blkName As String = "Pipe_" & StructObj.Name.Trim & "_" & hg
                                Using acBlkTblRecTemp As New BlockTableRecord
                                    If Not acBlkTbl.Has(blkName) Then
                                        acBlkTblRecTemp.Name = blkName
                                        acBlkTblRecTemp.Origin = StructObj.Position
                                        acBlkTblRecTemp.AppendEntity(Solid3dObj)
                                        acBlkTbl.UpgradeOpen()
                                        acBlkTbl.Add(acBlkTblRecTemp)
                                        tr.AddNewlyCreatedDBObject(acBlkTblRecTemp, True)
                                    End If
                                    Dim br As BlockReference = New BlockReference(StructObj.Position, acBlkTblRecTemp.Id)
                                    Dim btr As BlockTableRecord = tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                                    br.Layer = NameLayer
                                    br.ScaleFactors = New Scale3d(1, 1, 1)
                                    br.Rotation = 0
                                    btr.AppendEntity(br)
                                    tr.AddNewlyCreatedDBObject(br, True)
                                    '=======================================================================================
                                    'добавляем атрибуты
                                    For j As Integer = 0 To ArrayName.GetUpperBound(1)
                                        Dim Name As String = ArrayName(0, j)
                                        Dim Desk As String = ArrayName(1, j)
                                        Dim val As String = ArrayExcel(j, i)
                                        Dim ad As AttributeDefinition = New AttributeDefinition
                                        ad.SetDatabaseDefaults()
                                        Dim ar As AttributeReference = New AttributeReference()
                                        ar.SetAttributeFromBlock(ad, br.BlockTransform)
                                        br.AttributeCollection.AppendAttribute(ar)
                                        tr.AddNewlyCreatedDBObject(ar, True)
                                        ar.Layer = NameLayer
                                        ar.Tag = Name
                                        ar.TextString = val
                                        ar.Oblique = 0
                                        ar.Invisible = False
                                        ad.Prompt = Desk
                                    Next j
                                End Using
                            Next i
                        End If

                    ElseIf comboBox.Text Like "Изменить значение" Then
                        '===============================================================================
                        Dim TextBlock1 As TextBlock = button1.Content
                        If TextBlock1.Text Like "Колодец" Then
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                Dim hg As String = ArrayExcel(CountArrayName, i)
                                Dim IDObj As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                Dim StructObj As [Structure] = tr.GetObject(IDObj, OpenMode.ForWrite)
                                '=======================================================================================
                                For j As Integer = 0 To ArrayName.GetUpperBound(1)
                                    Dim Name As String = ArrayName(0, j)
                                    Dim Desk As String = ArrayName(1, j)
                                    Dim val As String = ArrayName(2, j)
                                    'имя колодца 0
                                    Try
                                        If Name Like "Name" Then
                                            StructObj.Name = val
                                        End If
                                        Dim newName As String = StructObj.Name

                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '1
                                    Try
                                        If Name Like "Desk" Then
                                            StructObj.Description = val
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '2
                                    Try
                                        If Name Like "Northing" Then
                                            StructObj.Northing = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '3
                                    Try
                                        If Name Like "Easting" Then
                                            StructObj.Easting = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '4
                                    Try
                                        If Name Like "ElevTop" Then
                                            StructObj.RimElevation = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '5
                                    Try
                                        If Name Like "ElevDown" Then
                                            StructObj.SumpElevation = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '6
                                    'Try
                                    '    If Name Like "PartDesk" Then
                                    '        StructObj.PartDescription = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '7
                                    'Try
                                    '    If Name Like "PartSizeNAme" Then
                                    '        StructObj.PartSizeName = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '8
                                    'Try
                                    '    If Name Like "PartSubType" Then
                                    '        StructObj.PartSubType = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '9
                                    'Try
                                    '    If Name Like "Diameter" Then
                                    '        StructObj.DiameterOrWidth = CDbl(val)
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '10
                                    'Try
                                    '    If Name Like "Lenght" Then
                                    '        StructObj.Length = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '11
                                    Try
                                        If Name Like "RimToSumpHeight" Then
                                            StructObj.RimToSumpHeight = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '12
                                    'Try
                                    '    If Name Like "Height" Then
                                    '        StructObj.Height = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '13
                                    'Try
                                    '    If Name Like "BarrelPipeClearance" Then
                                    '        StructObj.BarrelPipeClearance = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '14
                                    Try
                                        If Name Like "SurfElevAtInsPoint" Then
                                            StructObj.SurfaceAdjustmentValue = CDbl(val)
                                        End If
                                    Catch ex As CivilException
                                    Catch ex As ArgumentException
                                    End Try
                                    '15
                                    'Try
                                    '    If Name Like "ConeHeight" Then
                                    '        StructObj.ConeHeight = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '16
                                    'Try
                                    '    If Name Like "ConnectedPartCount" Then
                                    '        StructObj.ConnectedPartCount = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '17
                                    Try
                                        If Name Like "SurfaceAdjustmentValue" Then
                                            StructObj.SurfaceAdjustmentValue = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException

                                    End Try
                                    '18
                                    'Try
                                    '    If Name Like "ConnectedPipesCount" Then
                                    '        StructObj.ConnectedPipesCount = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '19
                                    'Try
                                    '    If Name Like "Domain" Then
                                    '        StructObj.Domain = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '20
                                    'Try
                                    '    If Name Like "FloorThickness" Then
                                    '        StructObj.FloorThickness = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '21
                                    'Try
                                    '    If Name Like "RuleSetStyleName" Then
                                    '        StructObj.Name = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '22
                                    'Try
                                    '    If Name Like "FrameDiameter" Then
                                    '        StructObj.FrameDiameter = val
                                    '    End If

                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '23
                                    'Try
                                    '    If Name Like "FrameHeight" Then
                                    '        StructObj.FrameHeight = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '24
                                    'Try
                                    '    If Name Like "Grate" Then
                                    '        StructObj.Grate = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '25
                                    'Try
                                    '    If Name Like "HeadwallBaseThickness" Then
                                    '        StructObj.HeadwallBaseThickness = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '26
                                    'Try
                                    '    If Name Like "RefAlignmentName" Then
                                    '        StructObj.RefAlignmentName = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '27
                                    'Try
                                    '    If Name Like "HeadwallBaseWidth" Then
                                    '        StructObj.HeadwallBaseWidth = CDbl(val)
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '28
                                    'Try
                                    '    If Name Like "InnerDiameterOrWidth" Then
                                    '        StructObj.InnerDiameterOrWidth = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '29
                                    'Try
                                    '    If Name Like "InnerLength" Then
                                    '        StructObj.InnerLength = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '30
                                    'Try
                                    '    If Name Like "Offset" Then
                                    '        StructObj.Offset = CDbl(val)
                                    '    End If
                                    'Catch ex As CivilException
                                    'End Try
                                    '32
                                    'Try
                                    '    If Name Like "VerticalPipeClearance" Then
                                    '        StructObj.VerticalPipeClearance = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '33
                                    Try
                                        If Name Like "RimElevation" Then
                                            StructObj.RimElevation = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    Try
                                        If Name Like "Rotation" Then
                                            StructObj.Rotation = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '35
                                    'Try
                                    '    If Name Like "Station" Then
                                    '        StructObj.Station = val
                                    '    End If
                                    'Catch ex As CivilException
                                    'End Try
                                    '36
                                    Try
                                        If Name Like "SumpDepth" Then
                                            StructObj.SumpDepth = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'смотрим пользовательские свойства
                                    Dim oData As PartDataField() = StructObj.PartData.GetAllDataFields()
                                    For Each idp As PartDataField In oData
                                        Try
                                            Dim NameTemp As String = idp.Name
                                            If Name Like NameTemp Then
                                                Dim Record As PartDataRecord = StructObj.PartData
                                                idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
                                                    idp.Value = CDbl(val)
                                                Else
                                                    idp.Value = val
                                                End If
                                                StructObj.PartData = Record
                                            End If
                                        Catch ex As InvalidOperationException

                                        End Try
                                    Next
                                Next j
                            Next i
                            '=========================================================================================================
                            'меняем значение для трубы
                        Else
                            For i As Integer = 0 To ArrayExcel.GetUpperBound(1)
                                Dim hg As String = ArrayExcel(CountArrayName, i)
                                Dim IDObj As ObjectId = AcadFunction.FuncReturnObjectID(hg)
                                Dim StructObj As Pipe = tr.GetObject(IDObj, OpenMode.ForWrite)
                                '=======================================================================================
                                For j As Integer = 0 To ArrayName.GetUpperBound(1)
                                    Dim Name As String = ArrayName(0, j)
                                    Dim Desk As String = ArrayName(1, j)
                                    Dim val As String = ArrayName(2, j)
                                    'имя колодца 0
                                    Try
                                        If Name Like "Name" Then
                                            StructObj.Name = val
                                        End If
                                    Catch ex As InvalidOperationException
                                    Catch ex As ArgumentException
                                    End Try
                                    '1
                                    Try
                                        If Name Like "Desk" Then
                                            StructObj.Description = val
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '4
                                    Try
                                        If Name Like "StartNorthing" Then
                                            Dim pt3d As Point3d = StructObj.StartPoint
                                            Dim newPt3d As Point3d = New Point3d(pt3d.X, CDbl(val), pt3d.Z)
                                            StructObj.StartPoint = newPt3d
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '5
                                    Try
                                        If Name Like "StartEasting" Then
                                            Dim pt3d As Point3d = StructObj.StartPoint
                                            Dim newPt3d As Point3d = New Point3d(CDbl(val), pt3d.Y, pt3d.Z)
                                            StructObj.StartPoint = newPt3d
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '4
                                    Try
                                        If Name Like "EndNorthing" Then
                                            Dim pt3d As Point3d = StructObj.EndPoint
                                            Dim newPt3d As Point3d = New Point3d(pt3d.X, CDbl(val), pt3d.Z)
                                            StructObj.EndPoint = newPt3d
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '5
                                    Try
                                        If Name Like "EndEasting" Then
                                            Dim pt3d As Point3d = StructObj.EndPoint
                                            Dim newPt3d As Point3d = New Point3d(CDbl(val), pt3d.Y, pt3d.Z)
                                            StructObj.EndPoint = newPt3d
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '6
                                    Try
                                        If Name Like "StartElev" Then
                                            Dim pt3d As Point3d = StructObj.StartPoint
                                            Dim newPt3d As Point3d = New Point3d(pt3d.X, pt3d.Y, CDbl(val))
                                            StructObj.StartPoint = newPt3d
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '7
                                    Try
                                        If Name Like "EndElev" Then
                                            Dim pt3d As Point3d = StructObj.EndPoint
                                            Dim newPt3d As Point3d = New Point3d(pt3d.X, pt3d.Y, CDbl(val))
                                            StructObj.EndPoint = newPt3d
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '8
                                    'Try
                                    '    If Name Like "PartDesk" Then
                                    '        StructObj.PartDescription = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''9
                                    'Try
                                    '    If Name Like "PartSizeNAme" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.PartSizeName
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''10
                                    'Try
                                    '    If Name Like "PartSubType" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.PartSubType
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '10
                                    'Try
                                    '    If Name Like "WallThickness" Then
                                    '        StructObj.WallThickness = val
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "Bearing" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Bearing
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    'Try
                                    '    If Name Like "CrossSectionalShape" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.CrossSectionalShape.ToString
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "StartStructure" Then
                                    '        Dim startStructID As ObjectId = StructObj.StartStructureId
                                    '        Dim StartStruct As [Structure] = tr.GetObject(startStructID, OpenMode.ForRead)
                                    '        ArrayExcel(i, CountArrayExcel) = StartStruct.Name
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "EndStructure" Then
                                    '        Dim startStructID As ObjectId = StructObj.EndStructureId
                                    '        Dim StartStruct As [Structure] = tr.GetObject(startStructID, OpenMode.ForRead)
                                    '        ArrayExcel(i, CountArrayExcel) = StartStruct.Name
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    'Try
                                    '    If Name Like "CoverOfStartPoint" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.CoverOfStartPoint
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "MinimumCover" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.MinimumCover
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "MaximumCover" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.MaximumCover
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "CoverOfEndpoint" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.CoverOfEndpoint
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "Domain" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Domain
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    Try
                                        If Name Like "EnergyGradeLineUp" Then
                                            StructObj.EnergyGradeLineUp = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    Try
                                        If Name Like "EnergyGradeLineDown" Then
                                            StructObj.EnergyGradeLineDown = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    'Try
                                    '    If Name Like "FlowDirection" Then
                                    '        StructObj.FlowDirection = CDbl(val)
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "FlowDirectionMethod" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.FlowDirectionMethod.ToString
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    Try
                                        If Name Like "FlowRate" Then
                                            StructObj.FlowRate = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    'Try
                                    '    If Name Like "HoldOnResizeType" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.HoldOnResizeType.ToString
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    Try
                                        If Name Like "HydraulicGradeLineUp" Then
                                            StructObj.HydraulicGradeLineUp = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    Try
                                        If Name Like "HydraulicGradeLineDown" Then
                                            StructObj.HydraulicGradeLineDown = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    'Try
                                    '    If Name Like "InnerDiameterOrWidth" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.InnerDiameterOrWidth
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "InnerHeight" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.InnerHeight
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    Try
                                        If Name Like "JunctionLoss" Then
                                            StructObj.JunctionLoss = CDbl(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    'Try
                                    '    If Name Like "Length2DCenterToCenter" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Length2DCenterToCenter
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "Length3DCenterToCenter" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Length3DCenterToCenter
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    'Try
                                    '    If Name Like "Material" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Material
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    'Try
                                    '    If Name Like "OuterDiameterOrWidth" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.OuterDiameterOrWidth
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "OuterHeight" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.OuterHeight
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    ''34
                                    'Try
                                    '    If Name Like "Radius" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Radius
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '34
                                    Try
                                        If Name Like "ReturnPeriod" Then
                                            StructObj.ReturnPeriod = CInt(val)
                                        End If
                                    Catch ex As InvalidOperationException
                                    End Try
                                    '34
                                    'Try
                                    '    If Name Like "Slope" Then
                                    '        ArrayExcel(i, CountArrayExcel) = StructObj.Slope
                                    '    End If
                                    'Catch ex As InvalidOperationException
                                    'End Try
                                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                                    'смотрим пользовательские свойства
                                    Dim oData As PartDataField() = StructObj.PartData.GetAllDataFields()
                                    For Each idp As PartDataField In oData
                                        Try
                                            Dim NameTemp As String = idp.Name
                                            If Name Like NameTemp Then
                                                Dim Record As PartDataRecord = StructObj.PartData
                                                idp = Record.GetDataFieldBy(idp.Context, idp.Index)
                                                If idp.DataType = PartCatalogDataType.Double Or idp.DataType = PartCatalogDataType.Int Then
                                                    idp.Value = CDbl(val)
                                                Else
                                                    idp.Value = val
                                                End If
                                                StructObj.PartData = Record
                                            End If
                                        Catch ex As InvalidOperationException
                                        Catch ex As ArgumentException
                                        End Try
                                    Next
                                Next j
                            Next i
                        End If
                    End If
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If
            tr.Commit()
        End Using
        docLock.Dispose()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменить тип фильтра
    Private Sub ComboBoxFilter_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBoxFilter.DropDownClosed
        Dim StrFilter As String = ComboBoxFilter.Text
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем фильтр
        Dim ArrayField As String() = Nothing
        Dim CountArrayField As Integer = 0
        Dim NameFilter As String = ComboBoxFilter.Text
        Dim NAmefileF As String = ComboBoxFilter.Tag & NameFilter & ".ini"
        'читаем поля
        If IO.File.Exists(ComboBoxFilter.Tag & NameFilter & ".ini") = True Then
            Using sr As IO.StreamReader = New IO.StreamReader(NAmefileF)
                Dim line As String
                Do
                    line = sr.ReadLine
                    If line Is Nothing Then
                        Exit Do
                    End If
                    If Not (line Like "") Then
                        ReDim Preserve ArrayField(CountArrayField)
                        ArrayField(CountArrayField) = line
                        CountArrayField += 1
                    End If
                Loop
            End Using
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'получаем выбранный объект
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim m_Editor As Editor = docCol.MdiActiveDocument.Editor
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim handleStr As String = textBox.Tag
            Dim idObj As ObjectId = AcadFunction.FuncReturnObjectID(handleStr)
            If Not (idObj.IsNull) = True Then
                Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(idObj, OpenMode.ForRead)
                Dim boolPipeProperties As Boolean = FuncReadPropertiesPipe(obj, tr, ArrayField)
                If boolPipeProperties = True Then
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'вставляем данные
                    dataGrid.ItemsSource = custdata
                End If
            End If
        End Using
    End Sub


End Class

