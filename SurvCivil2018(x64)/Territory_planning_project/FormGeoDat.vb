Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Collections.Generic
Imports System
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Drawing.Graphics
Imports System.Drawing
Imports System.IO
Imports System.String

Public Class FormGeoDat
    Inherits System.Windows.Forms.Form
    Public BoolForm As Integer = 1
    Public Sub GeoDat()
        Dim GeoDat1 As FormGeoDat = New FormGeoDat()
        Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(GeoDat1)
    End Sub
    Public Sub New()
        MyBase.New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавьте все инициализирующие действия после вызова InitializeComponent().

    End Sub

    'отмена
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        BoolForm = 0
        Me.Hide()
    End Sub
    'применить
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        BoolForm = 1
        Me.Hide()
    End Sub
End Class
