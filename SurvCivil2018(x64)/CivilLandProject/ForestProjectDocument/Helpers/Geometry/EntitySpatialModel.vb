Imports System.Xml.Serialization

''' <summary>
''' Пространственные данные (EntitySpatialType).
''' Choice: внешний файл с геоданными ИЛИ координаты прямо в XML.
''' По нашему сценарию используем второй вариант — координаты в XML.
''' </summary>
Public Class EntitySpatialModel

    ' --- Вариант 1: внешний файл ---
    <XmlElement("externalFormat")>
    Public Property ExternalFormat As KindEModel

    <XmlElement("storage")>
    Public Property Storage As KindEModel

    <XmlElement("specialFiles")>
    Public Property SpecialFiles As SpecialFilesModel

    ' --- Вариант 2: координаты в XML ---
    <XmlElement("csCode")>
    Public Property CsCode As KindEModel

    <XmlElement("spatialsElements")>
    Public Property SpatialsElements As SpatialsElementsModel

End Class
