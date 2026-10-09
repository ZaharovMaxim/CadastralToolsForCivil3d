Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Civil.Runtime
Imports Autodesk.Civil.Settings
Imports Autodesk.Civil.DatabaseServices.Styles
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors

Public Class FuncFeatureLine
    '===================================================================================================
    'функция создает стиль характерной линии
    Public Shared Function FuncCreateFeatureLineStyleByArrayProperties(ByVal ArrayProperty As String(,), Optional PutchTemplate As String = "") As FeatureLineStyle
        FuncCreateFeatureLineStyleByArrayProperties = Nothing
        If IsNothing(ArrayProperty) = True Then
            Exit Function
        End If
        If IsArray(ArrayProperty) = False Then
            Exit Function
        End If
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim indColor As Integer = 255
            Dim NameLayer As String = "0"
            Dim NameTypeLine As String = "Continuous"
            Dim ScaleLine As Double = 1
            Dim WightLine As LineWeight = LineWeight.ByLayer
            Dim NameStyleFeatureLine As String = "0-Характерная линия"
            Dim DeskStyleFeatyreLine = "Характерная линия"
            For i As Integer = 0 To ArrayProperty.GetUpperBound(1)
                Dim NameField As String = ArrayProperty(0, i)
                Dim ValStr As String = ArrayProperty(1, i).Trim
                '======================================================================
                'назначаем цвет
                If NameField Like "Цвет" Then
                    Dim nameColor As String = ValStr
                    If nameColor Like "ПоСлою" Then
                        indColor = 256
                    ElseIf nameColor Like "ПоБлоку" Then
                        indColor = 0
                    ElseIf nameColor Like "ByLayer" Then
                        indColor = 256
                    ElseIf nameColor Like "ByBlock" Then
                        indColor = 0
                    Else
                        indColor = Val(ValStr)
                        If indColor < 0 Then
                            indColor = 0
                        ElseIf indColor > 256 Then
                            indColor = 256
                        End If
                    End If
                ElseIf NameField Like "Наименование слоя" Then
                    NameLayer = ValStr
                ElseIf NameField Like "Вес линии" Then
                    If ValStr Like "ПоСлою" Then
                        WightLine = LineWeight.ByLayer
                    ElseIf ValStr Like "ПоБлоку" Then
                        WightLine = LineWeight.ByBlock
                    ElseIf ValStr Like "ByLayer" Then
                        WightLine = LineWeight.ByLayer
                    ElseIf ValStr Like "ByBlock" Then
                        WightLine = LineWeight.ByBlock
                    Else
                        Try
                            WightLine = Val(ValStr)
                        Catch ex As ArgumentException
                        End Try
                    End If
                ElseIf NameField Like "Масштаб" Then
                    ScaleLine = Val(ValStr)
                    If ScaleLine <= 0 Then
                        ScaleLine = 1
                    End If
                ElseIf NameField Like "Имя объекта" Then
                    If ValStr.Length > 0 Then
                        NameStyleFeatureLine = ValStr
                        NameTypeLine = ValStr
                    End If
                ElseIf NameField Like "Описание объекта" Then
                    DeskStyleFeatyreLine = ValStr
                End If
            Next i
            '========================================================================================================
            'создаем слой
            Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
            If acLyrTbl.Has(NameLayer) = False Then
                'копируем слой
                If IO.File.Exists(PutchTemplate) = True Then
                    Dim boolInsLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplate, NameLayer)
                    If boolInsLayer = False Then
                        NameLayer = "0"
                    End If
                Else
                    NameLayer = "0"
                End If
            End If
            '===========================================================================================================
            'создаем тип линии
            Dim acLineTypTbl As LinetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead)
            If acLineTypTbl.Has(NameTypeLine) = False Then
                'копируем тип линии
                If IO.File.Exists(PutchTemplate) = True Then
                    Dim boolInsTypeLine As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplate, NameTypeLine)
                    If boolInsTypeLine = False Then
                        NameTypeLine = "Continuous"
                    End If
                Else
                    NameTypeLine = "Continuous"
                End If
            End If
            '=============================================================================================================
            'создаем стиль Характерной линии
            If CivilDoc.Styles.FeatureLineStyles.Contains(NameStyleFeatureLine) = False Then
                Dim IdStyleFeatureLine As ObjectId = CivilDoc.Styles.FeatureLineStyles.Add(NameTypeLine)
                If IdStyleFeatureLine.IsNull = False Then
                    Dim StyleFeatureLine As FeatureLineStyle = tr.GetObject(IdStyleFeatureLine, OpenMode.ForWrite)
                    StyleFeatureLine.Description = DeskStyleFeatyreLine
                    '=====================================================================================================
                    StyleFeatureLine.FeatureLineDisplayStylePlan.Color = Color.FromColorIndex(ColorMethod.ByAci, indColor)
                    StyleFeatureLine.FeatureLineDisplayStylePlan.Layer = NameLayer
                    StyleFeatureLine.FeatureLineDisplayStylePlan.Linetype = NameTypeLine
                    StyleFeatureLine.FeatureLineDisplayStylePlan.LinetypeScale = ScaleLine
                    StyleFeatureLine.FeatureLineDisplayStylePlan.Lineweight = WightLine

                    StyleFeatureLine.FeatureLineDisplayStyleModel.Color = Color.FromColorIndex(ColorMethod.ByAci, indColor)
                    StyleFeatureLine.FeatureLineDisplayStyleModel.Layer = NameLayer
                    StyleFeatureLine.FeatureLineDisplayStyleModel.Linetype = NameTypeLine
                    StyleFeatureLine.FeatureLineDisplayStyleModel.LinetypeScale = ScaleLine
                    StyleFeatureLine.FeatureLineDisplayStylePlan.Lineweight = WightLine

                    FuncCreateFeatureLineStyleByArrayProperties = StyleFeatureLine
                End If
            End If
            tr.Commit()
        End Using
    End Function

End Class
