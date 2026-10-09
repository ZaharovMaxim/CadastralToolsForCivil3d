Imports System.Xml.Serialization

''' <summary>
''' Сведения о погибших лесных насаждениях (LostForestPlantationsType)
''' </summary>
Public Class LostForestPlantationsModel

    <XmlElement("year")>
    Public Property Year As String

    <XmlElement("typeDamage")>
    Public Property TypeDamage As String

    <XmlElement("typeBurning")>
    Public Property TypeBurning As KindEModel

    <XmlElement("yearDamage")>
    Public Property YearDamage As String

    <XmlElement("degreeDamage")>
    Public Property DegreeDamage As String

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("weakenedStock")>
    Public Property WeakenedStock As ValueUnitMeasurementModel

    <XmlElement("shrinkingStock")>
    Public Property ShrinkingStock As ValueUnitMeasurementModel

    <XmlElement("typeAccom")>
    Public Property TypeAccom As String

End Class
