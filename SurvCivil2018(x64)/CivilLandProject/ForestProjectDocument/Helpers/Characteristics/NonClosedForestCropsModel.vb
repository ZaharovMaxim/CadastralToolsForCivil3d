Imports System.Xml.Serialization

''' <summary>
''' Сведения о несомкнувшихся лесных культурах (NonClosedForestCropsType)
''' </summary>
Public Class NonClosedForestCropsModel

    <XmlElement("composition")>
    Public Property Composition As String

    <XmlElement("soilPrepareMethod")>
    Public Property SoilPrepareMethod As String

    <XmlElement("survivalRate")>
    Public Property SurvivalRate As Decimal?

    <XmlElement("treeAmount")>
    Public Property TreeAmount As Decimal?

    <XmlElement("treeQualityScore")>
    Public Property TreeQualityScore As Decimal?

    <XmlElement("poorStatusReason")>
    Public Property PoorStatusReason As String

    <XmlElement("cropDeathReason")>
    Public Property CropDeathReason As String

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("shrubbySpecies")>
    Public Property ShrubbySpecies As SpeciesModel

    <XmlElement("nonClosedForestType")>
    Public Property NonClosedForestType As String

    <XmlElement("addInfo")>
    Public Property AddInfo As String

End Class
