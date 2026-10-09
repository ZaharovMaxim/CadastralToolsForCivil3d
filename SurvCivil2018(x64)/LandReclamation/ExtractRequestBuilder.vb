Imports System.Windows.Forms
Imports System.Xml.Linq
Public Class ExtractRequestBuilder
    Public Property ResultData As Dictionary(Of String, Dictionary(Of String, String))
    Public Shared Function Build(data As Dictionary(Of String, String)) As XDocument
        Dim ns = XmlNamespaces.ProvisionChanges

        Dim root As New XElement(ns + "ForestObjectNotificationChangeRequest",
            New XAttribute(XNamespace.Xmlns + "ns", ns.NamespaceName)
        )

        ' === ForestPart ===
        If HasAny(data, {"ForestPart.numberForestPart", "ForestPart.cadNumberForestPart"}) Then
            Dim forestPart As New XElement(ns + "ForestPart")
            AddIf(forestPart, ns, "Action", GetVal(data, "ForestPart.Action"))
            AddIf(forestPart, ns, "numberForestPart", GetVal(data, "ForestPart.numberForestPart"))
            AddIf(forestPart, ns, "cadNumberForestPart", GetVal(data, "ForestPart.cadNumberForestPart"))
            root.Add(forestPart)
        End If

        ' === OKS ===
        If HasAny(data, {"OKS.numberOKS", "OKS.cadNumberOKS", "OKS.numberLandPlot"}) Then
            Dim oks As New XElement(ns + "OKS")
            AddIf(oks, ns, "Action", GetVal(data, "OKS.Action"))
            AddIf(oks, ns, "numberOKS", GetVal(data, "OKS.numberOKS"))
            AddIf(oks, ns, "cadNumberOKS", GetVal(data, "OKS.cadNumberOKS"))
            AddIf(oks, ns, "numberForestry", GetVal(data, "OKS.numberForestry"))
            AddIf(oks, ns, "numberLandPlot", GetVal(data, "OKS.numberLandPlot"))
            root.Add(oks)
        End If

        ' === Forestry ===
        If HasAny(data, {"Forestry.numberForestry", "Forestry.registryNumber"}) Then
            Dim forestry As New XElement(ns + "Forestry")
            AddIf(forestry, ns, "Action", GetVal(data, "Forestry.Action"))
            AddIf(forestry, ns, "numberForestry", GetVal(data, "Forestry.numberForestry"))
            AddIf(forestry, ns, "registryNumber", GetVal(data, "Forestry.registryNumber"))
            root.Add(forestry)
        End If

        ' === SpecialZone ===
        If HasAny(data, {"SpecialZone.registryNumber"}) Then
            Dim sz As New XElement(ns + "SpecialZone")
            AddIf(sz, ns, "Action", GetVal(data, "SpecialZone.Action"))
            AddIf(sz, ns, "registryNumber", GetVal(data, "SpecialZone.registryNumber"))
            root.Add(sz)
        End If

        Return New XDocument(New XDeclaration("1.0", "utf-8", Nothing), root)
    End Function

    Private Shared Function GetVal(data As Dictionary(Of String, String), key As String) As String
        If data.ContainsKey(key) Then Return data(key)
        Return Nothing
    End Function

    Private Shared Function HasAny(data As Dictionary(Of String, String), keys As String()) As Boolean
        For Each k In keys
            If data.ContainsKey(k) AndAlso Not String.IsNullOrWhiteSpace(data(k)) Then Return True
        Next
        Return False
    End Function

    Private Shared Sub AddIf(parent As XElement, ns As XNamespace, name As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        parent.Add(New XElement(ns + name, value))
    End Sub


End Class
