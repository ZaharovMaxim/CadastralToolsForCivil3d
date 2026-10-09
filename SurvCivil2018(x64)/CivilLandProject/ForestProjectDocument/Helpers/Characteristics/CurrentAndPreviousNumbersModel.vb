Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Ранее присвоенные номера (CurrentAndPreviousNumbersType)
''' </summary>
Public Class CurrentAndPreviousNumbersModel

    <XmlElement("previousNumbers")>
    Public Property PreviousNumbers As PreviousNumbersModel

    <XmlElement("currentNumberInfo")>
    Public Property CurrentNumberInfo As CurrentNumberInfoModel

End Class

''' <summary>
''' Ранее присвоенные номера (PreviousNumbersType)
''' </summary>
Public Class PreviousNumbersModel

    <XmlElement("previousNumberInfo")>
    Public Property PreviousNumberInfos As List(Of PreviousNumberModel)

End Class

''' <summary>
''' Ранее присвоенный номер (PreviousNumberType)
''' </summary>
Public Class PreviousNumberModel

    <XmlElement("previousNumber")>
    Public Property PreviousNumber As String

    <XmlElement("dateAssignmentNumber")>
    Public Property DateAssignmentNumber As String

    <XmlElement("assignmentOrganizationName")>
    Public Property AssignmentOrganizationName As String

End Class

''' <summary>
''' Последний присвоенный номер (CurrentNumberInfoType)
''' </summary>
Public Class CurrentNumberInfoModel

    <XmlElement("currentNumber")>
    Public Property CurrentNumber As String

    <XmlElement("dateAssignmentNumber")>
    Public Property DateAssignmentNumber As String

    <XmlElement("assignmentOrganizationName")>
    Public Property AssignmentOrganizationName As String

End Class
