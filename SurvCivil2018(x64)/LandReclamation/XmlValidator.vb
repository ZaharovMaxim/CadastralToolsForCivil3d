Imports System.IO
Imports System.Xml
Imports System.Xml.Schema
Imports Autodesk.AutoCAD.EditorInput
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class XmlValidator
    Public Shared Function LoadSchemasForChangeRequest(schemaDir As String) As XmlSchemaSet
        Dim schemaSet As New XmlSchemaSet()
        schemaSet.XmlResolver = New XmlUrlResolver()

        Dim fileName = "provisionChangesForestObjects_v1.1.0.xsd"
        Dim filePath = Path.Combine(schemaDir, fileName)

        If Not File.Exists(filePath) Then
            Throw New FileNotFoundException($"XSD не найдена: {filePath}")
        End If

        Using reader As XmlReader = XmlReader.Create(filePath)
            Dim schema As XmlSchema = XmlSchema.Read(reader, Nothing)
            schemaSet.Add(schema)
        End Using

        schemaSet.Compile()
        Return schemaSet
    End Function




    Public Shared Function LoadSchemas(schemaDir As String) As XmlSchemaSet
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim schemaSet As New XmlSchemaSet()
        schemaSet.XmlResolver = New XmlUrlResolver()

        If Not Directory.Exists(schemaDir) Then
            Throw New DirectoryNotFoundException($"Папка со схемами не найдена: {schemaDir}")
        End If

        ' === ЯВНЫЙ СПИСОК файлов, которые нужны для FORESTREQUESTXML ===
        ' Имена взяты из <xs:import schemaLocation="..."> в stateForestRegisterExtractTypes.xsd
        Dim filesToLoad As String() = {
        "simpleTypes.xsd",
        "complexTypes.xsd",
        "addressTypes.xsd",
        "subjectTypes.xsd",
        "popdContoursLocation_v0.2.0.xsd",
        "dUnitMeasurementKindTypes.xsd",
        "dForestResoursKindTypes.xsd",
        "dTreeNewKindTypes.xsd",
        "dOkpd2KindTypes.xsd",
        "dUsageTypeKindTypes.xsd",
        "dCategoryLandForestryKindTypes.xsd",
        "dTaxationCategoryKindTypes.xsd",
        "dPurposeForestsKindTypes.xsd",
        "dNameProtectedForestsKindTypes.xsd",
        "dCategoryLandKindTypes.xsd",
        "dLandTypeKindTypes.xsd",
        "dTypeForestKindTypes.xsd",
        "dTypeForestConditionsKindTypes.xsd",
        "dStateForestPlantKindTypes.xsd",
        "dAgeClassKindTypes.xsd",
        "dAgeGroupKindTypes.xsd",
        "dBonitetKindTypes.xsd",
        "dMarketabilityClassKindTypes.xsd",
        "dCuttingTypeKindTypes.xsd",
        "dTypeForestFireKindTypes.xsd",
        "dTypeHaymakingKindTypes.xsd",
        "dLocationTypeKindTypes.xsd",
        "dGardenTypeKindTypes.xsd",
        "dTypeGardenPlantsKindTypes.xsd",
        "dSwampTypeKindTypes.xsd",
        "dTypeVegetationKindTypes.xsd",
        "dSlopeSteepnessKindTypes.xsd",
        "dSlopeExposureKindTypes.xsd",
        "dErosionTypeKindTypes.xsd",
        "dErosionDegreeKindTypes.xsd",
        "dContaminationLandsKindTypes.xsd",
        "dMeasuresOZVLKindTypes.xsd",
        "attachmentTypes.xsd",
        "dFellingForestKindTypes.xsd",
        "dCuttingGroupCategoryKindTypes.xsd",
        "dTnVedWoodKindTypes.xsd",
        "dConstituentEntityKindTypes.xsd",
        "stateForestRegisterExtractTypes.xsd"
    }

        Console.WriteLine($"Загружаем XSD по явному списку: {filesToLoad.Length} файлов")

        For Each fileName In filesToLoad
            Dim fullPath = Path.Combine(schemaDir, fileName)
            If Not File.Exists(fullPath) Then
                ed.WriteMessage($"  [НЕ НАЙДЕН] {fileName}")
                Continue For
            End If

            Try
                Using reader As XmlReader = XmlReader.Create(fullPath)
                    Dim schema As XmlSchema = XmlSchema.Read(reader, Nothing)
                    schemaSet.Add(schema)
                    ed.WriteMessage($"  [OK] {fileName}")
                End Using
            Catch ex As Exception
                ed.WriteMessage($"  [ОШИБКА] {fileName}: {ex.Message}")
            End Try
        Next

        schemaSet.Compile()
        ed.WriteMessage("Схемы скомпилированы успешно.")
        Return schemaSet
    End Function
    Public Shared Function LoadSchemas3(schemaDir As String) As XmlSchemaSet
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim schemaSet As New XmlSchemaSet()
        schemaSet.XmlResolver = New XmlUrlResolver()
        If Not Directory.Exists(schemaDir) Then
            Throw New DirectoryNotFoundException($"Папка со схемами не найдена: {schemaDir}")
        End If

        Dim files = Directory.GetFiles(schemaDir, "*.xsd", SearchOption.AllDirectories)
        Console.WriteLine($"Найдено XSD-файлов: {files.Length}")

        ' Сначала соберём все namespace и посмотрим, какие дублируются
        Dim byNs As New Dictionary(Of String, List(Of String))()
        Dim parsed As New List(Of Tuple(Of String, XmlSchema))()

        For Each file In files
            Try
                Using reader As XmlReader = XmlReader.Create(file)
                    Dim schema As XmlSchema = XmlSchema.Read(reader, Nothing)
                    Dim ns = If(schema.TargetNamespace, "")
                    If Not byNs.ContainsKey(ns) Then byNs(ns) = New List(Of String)()
                    byNs(ns).Add(Path.GetFileName(file))
                    parsed.Add(Tuple.Create(file, schema))
                End Using
            Catch ex As Exception
                ed.WriteMessage($"  [не прочитан] {Path.GetFileName(file)}: {ex.Message}" & vbLf)
            End Try
        Next

        ' Показываем namespace'ы, у которых больше одного файла
        ed.WriteMessage("Namespace'ы с несколькими файлами:" & vbLf)
        For Each kv In byNs
            If kv.Value.Count > 1 Then
                Console.WriteLine($"  {kv.Key}")
                For Each f In kv.Value
                    Console.WriteLine($"      {f}")
                Next
            End If
        Next
        Console.WriteLine()

        ' Теперь добавляем в схему — по одному файлу на namespace
        Dim addedNs As New HashSet(Of String)()
        For Each pair In parsed
            Dim file = pair.Item1
            Dim schema = pair.Item2
            Dim ns = If(schema.TargetNamespace, "")
            If addedNs.Contains(ns) Then
                ed.WriteMessage($"  [пропущен дубль] {Path.GetFileName(file)} (namespace {ns} уже добавлен)" & vbLf)
                Continue For
            End If
            schemaSet.Add(schema)
            addedNs.Add(ns)
        Next

        schemaSet.Compile()
        ed.WriteMessage("Схемы скомпилированы успешно." & vbLf)
        Return schemaSet
    End Function
    Public Shared Function LoadSchemas2(schemaDir As String) As XmlSchemaSet
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim schemaSet As New XmlSchemaSet()
        schemaSet.XmlResolver = New XmlUrlResolver()

        If Not Directory.Exists(schemaDir) Then
            Throw New DirectoryNotFoundException($"Папка со схемами не найдена: {schemaDir}")
        End If

        Dim files = Directory.GetFiles(schemaDir, "*.xsd", SearchOption.AllDirectories)
        ed.WriteMessage($"Найдено XSD-файлов: {files.Length}")

        For Each file In files
            Try
                Using reader As XmlReader = XmlReader.Create(file)
                    Dim schema As XmlSchema = XmlSchema.Read(reader, Nothing)
                    schemaSet.Add(schema)
                End Using
            Catch ex As Exception
                ed.WriteMessage($"  [пропущен] {Path.GetFileName(file)}: {ex.Message}")
            End Try
        Next

        Try
            schemaSet.Compile()
            ed.WriteMessage("Схемы скомпилированы успешно.")
        Catch ex As Exception
            ed.WriteMessage($"Ошибка компиляции схем: {ex.Message}")
            Throw
        End Try

        Return schemaSet
    End Function

    ''' <summary>
    ''' Валидирует XML по заранее загруженному набору схем.
    ''' Возвращает список сообщений об ошибках.
    ''' </summary>
    Public Shared Function Validate(xmlPath As String,
                                schemas As XmlSchemaSet) As List(Of String)

        Dim errors As New List(Of String)

        Dim settings As New XmlReaderSettings With {
        .ValidationType = ValidationType.Schema,
        .Schemas = schemas,
        .ValidationFlags = XmlSchemaValidationFlags.ReportValidationWarnings
    }

        AddHandler settings.ValidationEventHandler,
        Sub(sender, e)
            Dim se = DirectCast(e, ValidationEventArgs)
            errors.Add($"{se.Severity}: {se.Message} (строка {se.Exception.LineNumber}, поз. {se.Exception.LinePosition})")
        End Sub

        Using reader As XmlReader = XmlReader.Create(xmlPath, settings)
            While reader.Read()
            End While
        End Using

        Return errors
    End Function
End Class
