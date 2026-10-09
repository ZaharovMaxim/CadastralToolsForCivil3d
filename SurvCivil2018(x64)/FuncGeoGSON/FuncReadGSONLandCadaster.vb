Imports System.IO
Imports Newtonsoft.Json
Public Class FuncReadGSONLandCadaster
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'IPI ROSRESTR
    'функция читает определенные в массиве поля и их значения в файле JSON для земельного участка (код 1)
    Public Shared Function FuncReadJsonParcels(Of T)(ByVal JSonStr As String, ByRef data As RosReestr_Parcels.Rootobject) As Boolean
        FuncReadJsonParcels = False
        Dim txtReader As TextReader = New StringReader(JSonStr)
        Using JReader As JsonReader = New JsonTextReader(txtReader)
            Dim s = New Newtonsoft.Json.JsonSerializer()
            data = s.Deserialize(Of RosReestr_Parcels.Rootobject)(JReader)
            FuncReadJsonParcels = True
        End Using
    End Function
    'функция читает определенные в массиве поля и их значения в файле JSON для строения (код 5)
    Public Shared Function FuncReadJsonBuild(Of T)(ByVal JSonStr As String, ByRef data As T) As Boolean
        FuncReadJsonBuild = False
        Dim txtReader As TextReader = New StringReader(JSonStr)
        Using JReader As JsonReader = New JsonTextReader(txtReader)
            Dim s = New Newtonsoft.Json.JsonSerializer()
            data = s.Deserialize(Of T)(JReader)
            FuncReadJsonBuild = True
        End Using
    End Function
    Public Shared Function FuncReadJsonText(ByVal JSonStr As String, ByRef Data As RosReestr) As Boolean
        FuncReadJsonText = False
        Dim txtReader As TextReader = New StringReader(JSonStr)
        Using JReader As JsonReader = New JsonTextReader(txtReader)
            Dim s = New Newtonsoft.Json.JsonSerializer()
            Data = s.Deserialize(Of RosReestr)(JReader)
            FuncReadJsonText = True
        End Using
    End Function

    Public Shared Function FuncReadJsonText2(Of T)(ByVal JSonStr As String, ByRef data As T) As Boolean
        FuncReadJsonText2 = False
        Dim txtReader As TextReader = New StringReader(JSonStr)
        Using JReader As JsonReader = New JsonTextReader(txtReader)
            Dim s = New Newtonsoft.Json.JsonSerializer()
            data = s.Deserialize(Of T)(JReader)
            FuncReadJsonText2 = True
        End Using
    End Function

    Public Shared Function FuncReadJsonText3(ByVal JSonStr As String, ArrayFields As String(), ByRef Arraydata As String(,)) As Boolean
        FuncReadJsonText3 = False
        Dim CountArrayData As Integer = 0
        Dim Razd As String = ", " & Chr(34)
        Dim TempArray As String() = JSonStr.Split(Razd)
        If IsArray(TempArray) = True Then
            Dim ArrayTemp2 As String() = Nothing
            Dim CountArrayTemp2 As Integer = 0
            For i As Integer = 0 To TempArray.Length - 1
                Dim tempDAta As String = TempArray(i)
                Dim pos As Integer = tempDAta.IndexOf(":")
                If pos > -1 Then
                    ReDim Preserve ArrayTemp2(CountArrayTemp2)
                    ArrayTemp2(CountArrayTemp2) = tempDAta.Trim
                    CountArrayTemp2 += 1
                Else
                    If CountArrayTemp2 > 0 Then
                        ArrayTemp2(CountArrayTemp2 - 1) = ArrayTemp2(CountArrayTemp2 - 1) & tempDAta
                    End If
                End If
            Next
            For i As Integer = 0 To ArrayFields.Length - 1
                Dim TempFields As String = ArrayFields(i)
                If TempFields.Trim.Length = 0 Then
                    Continue For
                End If
                For j As Integer = 0 To ArrayTemp2.Length - 1
                    Dim tempDAta As String = ArrayTemp2(j)
                    Dim str As String() = tempDAta.Split(":")
                    If IsArray(str) = True Then
                        If str.Length > 0 Then
                            Dim pos1 As Integer = str(0).IndexOf(Chr(34) & TempFields & Chr(34))
                            If pos1 > -1 Then
                                ReDim Preserve Arraydata(1, CountArrayData)
                                Arraydata(0, CountArrayData) = TempFields
                                Arraydata(1, CountArrayData) = str(1)
                                If str.Length > 1 Then
                                    For k As Integer = 2 To str.Length - 1
                                        Arraydata(1, CountArrayData) = Arraydata(1, CountArrayData) & ":" & str(k)
                                    Next k
                                End If
                                Arraydata(1, CountArrayData) = Arraydata(1, CountArrayData).Trim
                                CountArrayData += 1
                            End If
                        End If
                    End If
                Next j
            Next i
        End If
    End Function
    Public Shared Function FuncDecodeJsonText(ByVal JSonStr As String) As Boolean
        Dim TempStr As String = ""
        Dim TempSumbArray As String() = JSonStr.Split("\u")
        If IsArray(TempSumbArray) = True Then
            For i As Integer = 0 To TempSumbArray.Length - 1
                Dim TempSumb As String = TempSumbArray(i)
                Dim tempSumb2 As String = Mid(TempSumb, 4)
                If tempSumb2 Like "0410" Then

                End If
            Next
        End If
    End Function
End Class
