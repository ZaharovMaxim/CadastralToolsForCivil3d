Imports System.Xml.Serialization

''' <summary>
''' Сведения о лице, выполнявшем работы (ExpertsEType)
''' </summary>
Public Class ExpertModel

    <XmlElement("FullName")>
    Public Property FullName As PersonNameModel

    <XmlElement("organization")>
    Public Property Organization As String

    <XmlElement("registrationNumber")>
    Public Property RegistrationNumber As String

End Class
