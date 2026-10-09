Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class ExplodeObjectDWGFiles
    <CommandMethod("ExpodeProxyObjectUser")>
    Public Sub ExpodeProxyObjectUser()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '1. получем доступ к пользовательскому каталогу
        'временный каталог для автосохранения таблицы
        Dim TempDir As String = FuncFiles.FuncGetAbsoluteFolderPathName("Выберите папку", 1)
        '=============================================================================================
        'читаем файлы dwg
        If IO.Directory.Exists(TempDir) = True Then
            Try
                'создаем дерево в TreeView
                For Each fileDWG As String In Directory.GetFiles(TempDir)
                    Dim ext As String = Path.GetExtension(fileDWG)
                    If ext Like ".dwg" Then
                        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Open(fileDWG, False)
                        Dim docs As DocumentCollection = Application.DocumentManager
                        Dim db As Database = acDoc.Database()
                        acDoc.SendStringToExecute("EXPLODEALLPROXY ", False, False, True)
                        acDoc.SendStringToExecute("REMOVEALLPROXY ", False, False, True)
                        'If acDoc.IsReadOnly Then
                        '    acDoc.CloseAndDiscard()
                        'Else
                        '    If docs.MdiActiveDocument <> acDoc Then
                        '        docs.MdiActiveDocument = acDoc
                        '    End If
                        '    Dim isModified As Integer = System.Convert.ToInt32(Application.GetSystemVariable("DBMOD"))
                        '    If isModified = 0 Then
                        '        acDoc.SendStringToExecute("^C^C ", False, False, True)
                        '        acDoc.CloseAndDiscard()
                        '    Else
                        '        acDoc.SendStringToExecute("^C^C ", False, False, True)
                        '        acDoc.CloseAndSave(acDoc.Name)
                        '    End If
                        'End If
                    End If
                Next
            Catch ex As IOException
            Catch ex As UnauthorizedAccessException
            Catch ex As ArgumentException
            End Try
        End If
    End Sub
End Class
