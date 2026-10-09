Imports System.Windows
Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.Civil.DatabaseServices
Imports Label = Autodesk.Civil.DatabaseServices.Label
Public Class UserPropertiesTopoSymbols
    'Public PutchTemplateDwg As String = "d:\Civil3DToolsUtility\FileResources\Sample\" 'пути для файла шаблонов
    Public Sub New()
        InitializeComponent()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        '1. заполняем список с поверхностями
        Dim ArrayTinSurface As String() = Nothing
        Dim CountArrayTinSurface As Integer = 0
        '=========================================================================================================
        'ищем поверхность
        Dim collSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
        Dim NameSurf As String = ""
        'If collSurf.Count > 0 Then
        '    Using acTrans As Transaction = db.TransactionManager.StartTransaction()
        '        Dim Sq As Double = 0
        '        For Each idSurf As ObjectId In collSurf
        '            Try
        '                Dim acSurfTemp As TinSurface = acTrans.GetObject(idSurf, OpenMode.ForRead)
        '                ReDim Preserve ArrayTinSurface(CountArrayTinSurface)
        '                ArrayTinSurface(CountArrayTinSurface) = acSurfTemp.Name
        '                CountArrayTinSurface += 1
        '                If acSurfTemp.GetTerrainProperties.SurfaceArea2D > Sq Then
        '                    NameSurf = acSurfTemp.Name
        '                End If
        '            Catch ex As System.Exception
        '            End Try
        '        Next
        '        If IsArray(ArrayTinSurface) = True Then
        '            ComboSurface.ItemsSource = ArrayTinSurface
        '            ComboSurface.Text = NameSurf
        '        End If
        '    End Using
        'End If
        ''\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        ''1. заполняем список с трассами
        'Dim ArrayAlign As String() = Nothing
        'Dim CountArrayAlign As Integer = 0
        ''=========================================================================================================
        ''ищем трассы
        'Dim collAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
        'Dim NameAlign As String = ""
        'If collAlign.Count > 0 Then
        '    Using acTrans As Transaction = db.TransactionManager.StartTransaction()
        '        Dim Sq As Double = 0
        '        For Each idAlign As ObjectId In collAlign
        '            Try
        '                Dim acAlignTemp As Alignment = acTrans.GetObject(idAlign, OpenMode.ForRead)
        '                ReDim Preserve ArrayAlign(CountArrayAlign)
        '                ArrayAlign(CountArrayAlign) = acAlignTemp.Name
        '                CountArrayAlign += 1
        '                If acAlignTemp.Length > Sq Then
        '                    NameAlign = acAlignTemp.Name
        '                End If
        '            Catch ex As System.Exception
        '            End Try
        '        Next
        '        If IsArray(ArrayAlign) = True Then
        '            ComboAlign.ItemsSource = ArrayAlign
        '            ComboAlign.Text = NameAlign
        '        End If
        '    End Using
        'End If
    End Sub
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    'обновить трассы
    Private Sub ButtonUpdateAlign_Click(sender As Object, e As RoutedEventArgs) Handles ButtonUpdateAlign.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. заполняем список с трассами
        Dim ArrayAlign As String() = Nothing
        Dim CountArrayAlign As Integer = 0
        '=========================================================================================================
        'ищем трассы
        Dim collAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
        Dim NameAlign As String = ""
        If collAlign.Count > 0 Then
            Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                Dim Sq As Double = 0
                For Each idAlign As ObjectId In collAlign
                    Try
                        Dim acAlignTemp As Alignment = acTrans.GetObject(idAlign, OpenMode.ForRead)
                        ReDim Preserve ArrayAlign(CountArrayAlign)
                        ArrayAlign(CountArrayAlign) = acAlignTemp.Name
                        CountArrayAlign += 1
                        If acAlignTemp.Length > Sq Then
                            NameAlign = acAlignTemp.Name
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
                If IsArray(ArrayAlign) = True Then
                    ComboAlign.ItemsSource = ArrayAlign
                    ComboAlign.Text = NameAlign
                End If
            End Using
        End If
    End Sub

    Private Sub ButtonUpdateSurface_Click(sender As Object, e As RoutedEventArgs) Handles ButtonUpdateSurface.Click
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. заполняем список с поверхностями
        Dim ArrayTinSurface As String() = Nothing
        Dim CountArrayTinSurface As Integer = 0
        '=========================================================================================================
        'ищем поверхность
        Dim collSurf As ObjectIdCollection = CivilDoc.GetSurfaceIds
        Dim NameSurf As String = ""
        If collSurf.Count > 0 Then
            Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                Dim Sq As Double = 0
                For Each idSurf As ObjectId In collSurf
                    Try
                        Dim acSurfTemp As TinSurface = acTrans.GetObject(idSurf, OpenMode.ForRead)
                        ReDim Preserve ArrayTinSurface(CountArrayTinSurface)
                        ArrayTinSurface(CountArrayTinSurface) = acSurfTemp.Name
                        CountArrayTinSurface += 1
                        If acSurfTemp.GetTerrainProperties.SurfaceArea2D > Sq Then
                            NameSurf = acSurfTemp.Name
                        End If
                    Catch ex As System.Exception
                    End Try
                Next
                If IsArray(ArrayTinSurface) = True Then
                    ComboSurface.ItemsSource = ArrayTinSurface
                    ComboSurface.Text = NameSurf
                End If
            End Using
        End If
    End Sub

    Private Sub UserPropertiesTopoSymbols_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged

    End Sub
End Class
