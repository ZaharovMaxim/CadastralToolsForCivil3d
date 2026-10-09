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
Public Class CadastrData
    Public LandCadastrPalette As Autodesk.AutoCAD.Windows.PaletteSet = Nothing
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'нажата кнопка сделать запрос
    Private Sub Button_Click_1(sender As Object, e As Windows.RoutedEventArgs)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed1 As Editor = acDoc.Editor
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
        Try
            Dim docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
            Using docLock
                If Button_Zapros.Content Like "Получить список кварталов" Then
                    '1. Выбираем объект
                    Dim acSSPrompt As PromptEntityResult = ed1.GetEntity("^C^C " & vbLf & "Выберите полилинию или трассу: ")
                    If Not (acSSPrompt.Status = PromptStatus.OK) Then
                        Exit Sub
                    End If
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        Dim acEnt As Autodesk.AutoCAD.DatabaseServices.Entity = acTrans.GetObject(acSSPrompt.ObjectId, OpenMode.ForRead)





                    End Using


                End If
            End Using
        Catch ex As Exception
        End Try

    End Sub

End Class
