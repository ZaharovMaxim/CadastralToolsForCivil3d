Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.StreamWriter
Imports System.String
Imports System.IO.File
Imports System.Net
Imports System.Xml
Imports Microsoft.Office.Interop
Imports OSGeo.MapGuide
Imports System.Drawing 'для рисунков jpg, tiff
Imports System.Text
Imports Autodesk.Gis.Map.HostMapApplicationServices
Imports Autodesk.Gis.Map.ObjectData
Imports Autodesk.Gis.Map
Imports System.Globalization
Imports System.Windows.Media.Imaging

Imports System.Security
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports Newtonsoft.Json
Public Class LandCadastrPalette
    Public boolLandCadastrPallette As Boolean = False
    '=========================================================================================================
    'сведения о земельном участке
    '================================================================================================
    'росреестр
    <CommandMethod("InformRosreestrObj")>
    Public Sub InformRosreestrObj()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        'находим все директории с файлами
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        'ищем пути для таблицы XML
        If boolLandCadastrPallette = False Then
            Dim RosreestrPalette As CadastrData = New CadastrData
            Dim LandPal As PaletteSet = New Autodesk.AutoCAD.Windows.PaletteSet("Росреестр")
            '========================================================================================
            Dim UserSize As Drawing.Size = New Drawing.Size(300, 800)
            LandPal.Size = UserSize
            LandPal.DockEnabled = DockSides.Left + DockSides.Right
            LandPal.AddVisual("Росреестр", RosreestrPalette)
            LandPal.Visible = True
            'типы оюъектов
            Dim ArrayProp As String() = {"Нет данных", "Участки", "ОКС", "ЗОУИТ", "Тер.зона", "Квартал"}
            RosreestrPalette.Combo_Type_Object.ItemsSource = ArrayProp
            'действие
            Dim ArrayProp1 As String() = {"Получить список кварталов", "Получить сведения из ЕГРН", "Запросить растр"}
            RosreestrPalette.Combo_Type_d.ItemsSource = ArrayProp1
            boolLandCadastrPallette = True
            LandPal.Visible = True
        End If
        boolLandCadastrPallette = False
    End Sub
End Class
