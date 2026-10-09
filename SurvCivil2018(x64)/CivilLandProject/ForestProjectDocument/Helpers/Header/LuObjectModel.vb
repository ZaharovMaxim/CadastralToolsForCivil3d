Imports System.Xml.Serialization

''' <summary>
''' Объект лесоустройства (LuObjectEType).
''' Choice из трёх вариантов: лесничество / участковое лесничество / лесной участок.
''' Заполняется только один из трёх.
''' </summary>
Public Class LuObjectModel

    <XmlElement("forestryType")>
    Public Property Forestry As ForestryModel

    <XmlElement("steadForestryType")>
    Public Property SteadForestry As SteadForestryModel

    <XmlElement("areasForest")>
    Public Property AreasForest As String

End Class
