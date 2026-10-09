Imports System.Xml.Serialization

''' <summary>
''' Характеристики лесного квартала (CharacteristicsForestQuarterType)
''' </summary>
Public Class CharacteristicsForestQuarterModel

    <XmlElement("forestryName")>
    Public Property ForestryName As ForestryModel

    <XmlElement("steadForestryName")>
    Public Property SteadForestryName As SteadForestryModel

    <XmlElement("nameTract")>
    Public Property NameTract As String

    <XmlElement("fgislkForestQuarterNumber")>
    Public Property FgislkForestQuarterNumber As String

    <XmlElement("luForestQuarterNumber")>
    Public Property LuForestQuarterNumber As String

    <XmlElement("taxationCategory")>
    Public Property TaxationCategory As KindEModel

End Class
