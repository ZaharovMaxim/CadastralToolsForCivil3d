Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports CadastralTools
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

''' <summary>
''' Команды AutoCAD/Civil 3D для формирования XML-документов
''' лесной проектной документации ФГИС ЛК.
''' </summary>
Public Class CmdForestProjectDocumentation

    ''' <summary>
    ''' Команда OnliXML — диалог формирования XML ОНЛИ
    ''' (внесение сведений в ГЛР об объектах, не связанных
    ''' с созданием лесной инфраструктуры, созданных в границах лесничеств).
    ''' </summary>
    <CommandMethod("OnliXML")>
    Public Sub OnliXML()
        Try
            Using frm As New frmOnli()
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
            Application.ShowAlertDialog("Ошибка формирования XML ОНЛИ: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Команда EditTaxationXML — диалог формирования XML «Описание изменений
    ''' лесотаксационных выделов» (editTaxationByVlscnOzul).
    ''' </summary>
    <CommandMethod("EditTaxationXML")>
    Public Sub EditTaxationXML()
        Try
            Using frm As New frmEditTaxation()
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
            Application.ShowAlertDialog("Ошибка формирования XML описания изменений выделов: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Команда TaxationDescriptionXML — диалог формирования XML «Таксационное
    ''' описание лесотаксационных выделов» (taxationDescription v2.0.0).
    ''' Форма немодальная: нужна для выбора замкнутой полилинии контура на чертеже.
    ''' </summary>
    <CommandMethod("TaxationDescriptionXML")>
    Public Sub TaxationDescriptionXML()
        Try
            Dim frm As New frmTaxationDescription()
            ' Инжекция реализации выбора полилинии на чертеже (CadastralTools не ссылается на AutoCAD)
            frm.ContourPicker = AddressOf MdPickContour.PickClosedPolyline
            AddHandler frm.FormClosed, Sub(s As Object, ev As FormClosedEventArgs) frm.Dispose()
            Application.ShowModelessDialog(frm)
        Catch ex As Exception
            Application.ShowAlertDialog("Ошибка формирования таксационного описания: " & ex.Message)
        End Try
    End Sub

End Class
