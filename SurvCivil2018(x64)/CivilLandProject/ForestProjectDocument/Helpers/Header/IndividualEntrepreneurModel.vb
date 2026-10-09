Imports System.Xml.Serialization

''' <summary>
''' Сведения об индивидуальном предпринимателе (IndividualEntrepreneurEType)
''' </summary>
Public Class IndividualEntrepreneurModel

    <XmlElement("fio")>
    Public Property Fio As PersonNameModel

    <XmlElement("inn")>
    Public Property Inn As String

    <XmlElement("regDate")>
    Public Property RegDate As String

    <XmlElement("iDocument")>
    Public Property IDocument As String

    <XmlElement("location")>
    Public Property Location As LocationModel

    <XmlElement("snils")>
    Public Property Snils As String

End Class
