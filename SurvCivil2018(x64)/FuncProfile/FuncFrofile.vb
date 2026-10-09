Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Colors
Public Class FuncFrofile
    '====================================================================================================================================================
    'функция заносит в массив пикеты конца и начала элементов вертикальной геометрии
    Public Shared Function FuncArrayPKVertGeom(ByVal ProfLine As Profile, ByRef ArrayFrofPK As String(,), Optional ByVal ErrRound As Double = 3) As Boolean
        FuncArrayPKVertGeom = False
        Try
            Dim CountArrayProfPK As Integer = 0
            If IsNothing(ArrayFrofPK) = False Then
                CountArrayProfPK = ArrayFrofPK.GetUpperBound(1) + 1
            End If
            Dim EntCollection As ProfileEntityCollection = ProfLine.Entities
            'перебираем сегменты
            For Each myAe As ProfileEntity In EntCollection
                Dim ProfEnt As ProfileEntityType = myAe.EntityType

                Select Case ProfEnt
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'прямой участок
                    Case ProfileEntityType.Tangent
                        Dim PLine As ProfileTangent = myAe
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PLine.StartStation
                        ArrayFrofPK(1, CountArrayProfPK) = "StartTangent"
                        CountArrayProfPK += 1
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PLine.EndStation
                        ArrayFrofPK(1, CountArrayProfPK) = "EndTangent"
                        CountArrayProfPK += 1
                    Case ProfileEntityType.Circular
                        Dim PCirc As ProfileCircular = myAe
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PCirc.StartStation
                        ArrayFrofPK(1, CountArrayProfPK) = "StartCircular"
                        CountArrayProfPK += 1
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PCirc.EndStation
                        ArrayFrofPK(1, CountArrayProfPK) = "EndCircular"
                        CountArrayProfPK += 1
                    Case ProfileEntityType.ParabolaAsymmetric
                        Dim PParAsym As ProfileParabolaAsymmetric = myAe
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PParAsym.StartStation
                        ArrayFrofPK(1, CountArrayProfPK) = "StartParabolaAsymmetric"
                        CountArrayProfPK += 1
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PParAsym.EndStation
                        ArrayFrofPK(1, CountArrayProfPK) = "EndParabolaAsymmetric"
                        CountArrayProfPK += 1
                    Case ProfileEntityType.ParabolaSymmetric
                        Dim PParSym As ProfileParabolaSymmetric = myAe
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PParSym.StartStation
                        ArrayFrofPK(1, CountArrayProfPK) = "StartParabolaSymmetric"
                        CountArrayProfPK += 1
                        ReDim Preserve ArrayFrofPK(1, CountArrayProfPK)
                        ArrayFrofPK(0, CountArrayProfPK) = PParSym.EndStation
                        ArrayFrofPK(1, CountArrayProfPK) = "EndParabolaSymmetric"
                        CountArrayProfPK += 1
                End Select
            Next
            FuncArrayPKVertGeom = True
        Catch
        End Try
    End Function
End Class
