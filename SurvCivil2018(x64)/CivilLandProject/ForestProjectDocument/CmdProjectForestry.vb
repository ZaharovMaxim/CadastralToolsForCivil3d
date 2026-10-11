Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports CadastralTools
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

''' <summary>
''' Команды AutoCAD/Civil 3D для формирования XML «Проект лесничества»
''' (projectForestry v2.2.0) для ФГИС ЛК.
''' Геометрия в документе не заполняется (objectGuid — идентификатор
''' контура в json-файле), выбор примитивов не требуется.
''' </summary>
Public Class CmdProjectForestry

    ''' <summary>
    ''' Команда ProjectForestryXML — диалог формирования XML «Проект лесничества».
    ''' </summary>
    <CommandMethod("ProjectForestryXML")>
    Public Sub ProjectForestryXML()
        Try
            Using frm As New frmProjectForestry()
                ' Привязка диалога к главному окну AutoCAD для модального поведения
                Dim owner As New NativeWindow()
                owner.AssignHandle(Application.MainWindow.Handle)
                Try
                    frm.ShowDialog(owner)
                Finally
                    owner.ReleaseHandle()
                End Try
            End Using
        Catch ex As Exception
            Application.ShowAlertDialog("Ошибка формирования XML проекта лесничества: " & ex.Message)
        End Try
    End Sub

End Class
