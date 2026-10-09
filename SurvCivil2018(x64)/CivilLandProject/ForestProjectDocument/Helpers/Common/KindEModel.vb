Imports System.Xml.Serialization

''' <summary>
''' Универсальный тип "код + значение" для всех справочников KindEType.
''' Используется везде, где в XSD встречается пара code/value.
''' </summary>
<XmlType(Namespace:="http://rosleshoz.gov.ru/fgis-lk/common/simple/types/2.0")>
Public Class KindEModel

    ''' <summary>
    ''' Код справочника
    ''' </summary>
    <XmlElement("code")>
    Public Property Code As String

    ''' <summary>
    ''' Значение справочника
    ''' </summary>
    <XmlElement("value")>
    Public Property Value As String

End Class
