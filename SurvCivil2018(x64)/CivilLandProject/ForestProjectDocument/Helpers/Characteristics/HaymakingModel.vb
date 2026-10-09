Imports System.Xml.Serialization

''' <summary>
''' Сведения о сенокосах (HaymakingType)
''' </summary>
Public Class HaymakingModel

    <XmlElement("typeHaymaking")>
    Public Property TypeHaymaking As KindEModel

    <XmlElement("conditionHaymaking")>
    Public Property ConditionHaymaking As KindEModel

    <XmlElement("qualityHaymaking")>
    Public Property QualityHaymaking As KindEModel

    <XmlElement("naturalEnhanced")>
    Public Property NaturalEnhanced As KindEModel

    <XmlElement("degreeOvergrowth")>
    Public Property DegreeOvergrowth As Decimal?

    <XmlElement("convenienceHayHarvesting")>
    Public Property ConvenienceHayHarvesting As String

    <XmlElement("mechanizationWork")>
    Public Property MechanizationWork As String

    <XmlElement("averageYield")>
    Public Property AverageYield As ValueUnitMeasurementModel

    <XmlElement("erosionProcesses")>
    Public Property ErosionProcesses As String

    <XmlElement("speciesName")>
    Public Property SpeciesName As KindEModel

End Class
