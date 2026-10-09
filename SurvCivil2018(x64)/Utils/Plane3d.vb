Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Collections.Generic
Imports System
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.AutoCAD.LayerManager.LayerFilter
Imports Autodesk.AutoCAD.Windows.OpenFileDialog
Imports Autodesk.AutoCAD.Colors
Imports System.Drawing.Graphics
Imports System.Drawing
Imports System.IO
Imports System.String
' Класс для представления точки в трехмерном пространстве
Public Class Point3Dnew
    Public Property X As Double
    Public Property Y As Double
    Public Property Z As Double

    Public Sub New(x As Double, y As Double, z As Double)
        Me.X = x
        Me.Y = y
        Me.Z = z
    End Sub
End Class

' Класс для представления плоскости
Public Class PlaneNew
    Public Property Normal As Point3d
    Public Property D As Double

    Public Sub New(a As Double, b As Double, c As Double, d As Double)
        Me.Normal = New Point3d(a, b, c)
        Me.D = d
    End Sub
End Class


Public Class Plane3dnew
    ' Функция для проверки, лежат ли четыре точки в одной плоскости
    Function ArePointsCoplanar(P1 As Point3Dnew, P2 As Point3Dnew, P3 As Point3Dnew, P4 As Point3Dnew) As Boolean
        ' Векторы, лежащие на плоскости
        Dim v1 As New Point3Dnew(P2.X - P1.X, P2.Y - P1.Y, P2.Z - P1.Z)
        Dim v2 As New Point3Dnew(P3.X - P1.X, P3.Y - P1.Y, P3.Z - P1.Z)
        Dim v3 As New Point3Dnew(P4.X - P1.X, P4.Y - P1.Y, P4.Z - P1.Z)

        ' Смешанное произведение векторов (объем параллелепипеда)
        Dim volume As Double = v1.X * (v2.Y * v3.Z - v2.Z * v3.Y) -
                               v1.Y * (v2.X * v3.Z - v2.Z * v3.X) +
                               v1.Z * (v2.X * v3.Y - v2.Y * v3.X)

        ' Если объем равен нулю, точки лежат в одной плоскости
        Return Math.Abs(volume) < Double.Epsilon
    End Function

    ' Функция для построения уравнения плоскости по трем точкам
    Function GetPlaneFromPoints(P1 As Point3Dnew, P2 As Point3Dnew, P3 As Point3Dnew) As PlaneNew
        ' Векторы, лежащие на плоскости
        Dim v1 As New Point3Dnew(P2.X - P1.X, P2.Y - P1.Y, P2.Z - P1.Z)
        Dim v2 As New Point3Dnew(P3.X - P1.X, P3.Y - P1.Y, P3.Z - P1.Z)
        ' Вектор нормали к плоскости через векторное произведение
        Dim normal As New Point3Dnew(
            v1.Y * v2.Z - v1.Z * v2.Y,
            v1.Z * v2.X - v1.X * v2.Z,
            v1.X * v2.Y - v1.Y * v2.X
        )

        Dim D As Double = -(normal.X * P1.X + normal.Y * P1.Y + normal.Z * P1.Z)
        Return New PlaneNew(normal.X, normal.Y, normal.Z, D)
    End Function

    Function FindIntersection(A As Point3Dnew, B As Point3Dnew, plane As PlaneNew) As Point3Dnew
        ' Вектор направления отрезка
        Dim direction As New Point3d(B.X - A.X, B.Y - A.Y, B.Z - A.Z)

        ' Параметрическое уравнение отрезка: P = A + t * direction
        ' Подставляем в уравнение плоскости: plane.Normal.X * (A.X + t * direction.X) + ... + plane.D = 0
        ' Решаем уравнение относительно t

        Dim numerator As Double = plane.Normal.X * A.X + plane.Normal.Y * A.Y + plane.Normal.Z * A.Z + plane.D
        Dim denominator As Double = plane.Normal.X * direction.X + plane.Normal.Y * direction.Y + plane.Normal.Z * direction.Z

        ' Если знаменатель равен нулю, отрезок параллелен плоскости
        If Math.Abs(denominator) < Double.Epsilon Then
            Return Nothing
        End If

        Dim t As Double = -numerator / denominator

        ' Если t находится в диапазоне [0, 1], то точка пересечения лежит на отрезке
        If t >= 0 AndAlso t <= 1 Then
            Return New Point3Dnew(A.X + t * direction.X, A.Y + t * direction.Y, A.Z + t * direction.Z)
        Else
            Return Nothing
        End If
    End Function


    <CommandMethod("workPlane11111")>
    Public Sub writeVertexPline111111()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim pr1 As PromptPointResult = ed.GetPoint("введите точку 1 плоскости: ")
            Dim pr2 As PromptPointResult = ed.GetPoint("введите точку 2 плоскости: ")
            Dim pr3 As PromptPointResult = ed.GetPoint("введите точку 3 плоскости: ")
            Dim pr4 As PromptPointResult = ed.GetPoint("введите точку 4 плоскости: ")

            Dim p1 As Point3Dnew = New Point3Dnew(pr1.Value.X, pr1.Value.Y, pr1.Value.Z)
            Dim p2 As Point3Dnew = New Point3Dnew(pr2.Value.X, pr2.Value.Y, pr2.Value.Z)
            Dim p3 As Point3Dnew = New Point3Dnew(pr3.Value.X, pr3.Value.Y, pr3.Value.Z)
            Dim p4 As Point3Dnew = New Point3Dnew(pr4.Value.X, pr4.Value.Y, pr4.Value.Z)


            If ArePointsCoplanar(p1, p2, p3, p4) Then
                ' Строим уравнение плоскости по трем точкам
                Dim plane As PlaneNew = GetPlaneFromPoints(p1, p2, p3)
                Dim pr5 As PromptPointResult = ed.GetPoint("введите точку 1 отрезка: ")
                Dim pr6 As PromptPointResult = ed.GetPoint("введите точку 2 отрезка: ")

                Dim pA As Point3Dnew = New Point3Dnew(pr5.Value.X, pr5.Value.Y, pr5.Value.Z)
                Dim pB As Point3Dnew = New Point3Dnew(pr6.Value.X, pr6.Value.Y, pr6.Value.Z)

                ' Находим точку пересечения
                Dim intersectionPoint As Point3Dnew = FindIntersection(pA, pB, plane)
                If intersectionPoint IsNot Nothing Then
                    MsgBox("Точка пересечения: " & intersectionPoint.X & vbLf & intersectionPoint.Y & vbLf & intersectionPoint.Z)
                Else
                    Console.WriteLine("Отрезок и плоскость не пересекаются.")
                End If
            Else
                Console.WriteLine("Точки не лежат в одной плоскости.")
            End If

        End Using
    End Sub


    <CommandMethod("intersectTriangle")>
    Public Sub intersectTriangle()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim pr1 As PromptPointResult = ed.GetPoint("введите точку 1 плоскости: ")
            Dim pr2 As PromptPointResult = ed.GetPoint("введите точку 2 плоскости: ")
            Dim pr3 As PromptPointResult = ed.GetPoint("введите точку 3 плоскости: ")
            Dim pr4 As PromptPointResult = ed.GetPoint("введите точку 4 плоскости: ")

            Dim p1 As Vector3d = New Vector3d(pr1.Value.X, pr1.Value.Y, pr1.Value.Z)
            Dim p2 As Vector3d = New Vector3d(pr2.Value.X, pr2.Value.Y, pr2.Value.Z)
            Dim p3 As Vector3d = New Vector3d(pr3.Value.X, pr3.Value.Y, pr3.Value.Z)
            Dim p4 As Vector3d = New Vector3d(pr4.Value.X, pr4.Value.Y, pr4.Value.Z)

            Dim pr5 As PromptPointResult = ed.GetPoint("введите точку 1 отрезка: ")
            Dim pr6 As PromptPointResult = ed.GetPoint("введите точку 2 отрезка: ")

            Dim startPoint As Vector3d = New Vector3d(pr5.Value.X, pr5.Value.Y, pr5.Value.Z)
            Dim endPoint As Vector3d = New Vector3d(pr6.Value.X, pr6.Value.Y, pr6.Value.Z)

            Dim intersection As Vector3d = IntersectSegmentWithPlane(p1, p2, p3, p4, startPoint, endPoint)
            If intersection.X <> 0 And intersection.Y <> 0 And intersection.Z <> 0 Then
                MsgBox("Точка пересечения: " & intersection.X & vbLf & intersection.Y & vbLf & intersection.Z)
            End If
        End Using
    End Sub



    Public Shared Function IntersectSegmentWithPlane(
    ByVal planePoint1 As Vector3d,
    ByVal planePoint2 As Vector3d,
    ByVal planePoint3 As Vector3d,
    ByVal planePoint4 As Vector3d,
    ByVal segmentStart As Vector3d,
    ByVal segmentEnd As Vector3d) As Vector3d

        ' Проверка на вырожденность четырехугольника
        If planePoint1.Equals(planePoint2) OrElse
       planePoint1.Equals(planePoint3) OrElse
       planePoint1.Equals(planePoint4) Then
            Return Nothing
        End If

        ' 1. Определяем плоскость по трем точкам (первые три точки)
        Dim v1 As Vector3d = planePoint2 - planePoint1
        Dim v2 As Vector3d = planePoint3 - planePoint1

        ' Нормаль к плоскости
        Dim normal As Vector3d = v1.CrossProduct(v2)

        ' Проверка на вырожденность (точки коллинеарны)
        If normal.Length <= Tolerance.Global.EqualPoint Then
            Return Nothing
        End If

        normal = normal.GetNormal()

        ' 2. Находим точку пересечения прямой с плоскостью
        Dim segmentDir As Vector3d = segmentEnd - segmentStart
        Dim w0 As Vector3d = segmentStart - planePoint1

        ' Вычисляем параметры для точки пересечения
        Dim a As Double = -normal.DotProduct(w0)
        Dim b As Double = normal.DotProduct(segmentDir)

        ' Проверка на параллельность (b близко к нулю)
        If Math.Abs(b) < Tolerance.Global.EqualPoint Then
            ' Прямая параллельна плоскости или лежит в ней
            Return Nothing
        End If

        ' Вычисляем параметр t для точки пересечения
        Dim t As Double = a / b

        ' Проверка, что точка пересечения лежит на отрезке (t между 0 и 1)
        If t < 0 - Tolerance.Global.EqualPoint OrElse t > 1 + Tolerance.Global.EqualPoint Then
            Return Nothing ' Точка пересечения вне отрезка
        End If

        ' Вычисляем точку пересечения
        Dim intersectionPoint As Vector3d = segmentStart + segmentDir * t

        ' 3. Проверяем, лежит ли точка внутри четырехугольника
        If IsPointInQuadrilateral(intersectionPoint, planePoint1, planePoint2, planePoint3, planePoint4) Then
            Return intersectionPoint
        Else
            Return Nothing
        End If
    End Function

    Private Shared Function IsPointInQuadrilateral(
    ByVal point As Vector3d,
    ByVal p1 As Vector3d,
    ByVal p2 As Vector3d,
    ByVal p3 As Vector3d,
    ByVal p4 As Vector3d) As Boolean

        ' Проверяем, что все точки лежат в одной плоскости
        ' (упрощенно - считаем что они в одной плоскости по условию задачи)

        ' Метод разбиения на два треугольника и проверка принадлежности
        Return IsPointInTriangle(point, p1, p2, p3) OrElse
               IsPointInTriangle(point, p1, p3, p4)
    End Function

    Private Shared Function IsPointInTriangle(
        ByVal point As Vector3d,
        ByVal t1 As Vector3d,
        ByVal t2 As Vector3d,
        ByVal t3 As Vector3d) As Boolean

        ' Вычисляем барицентрические координаты
        Dim v0 As Vector3d = t3 - t1
        Dim v1 As Vector3d = t2 - t1
        Dim v2 As Vector3d = point - t1

        ' Вычисляем скалярные произведения
        Dim dot00 As Double = v0.DotProduct(v0)
        Dim dot01 As Double = v0.DotProduct(v1)
        Dim dot02 As Double = v0.DotProduct(v2)
        Dim dot11 As Double = v1.DotProduct(v1)
        Dim dot12 As Double = v1.DotProduct(v2)

        ' Вычисляем барицентрические координаты
        Dim invDenom As Double = 1.0 / (dot00 * dot11 - dot01 * dot01)
        Dim u As Double = (dot11 * dot02 - dot01 * dot12) * invDenom
        Dim v As Double = (dot00 * dot12 - dot01 * dot02) * invDenom

        ' Проверяем, лежит ли точка внутри треугольника
        Return (u >= -Tolerance.Global.EqualPoint) AndAlso
               (v >= -Tolerance.Global.EqualPoint) AndAlso
               (u + v <= 1.0 + Tolerance.Global.EqualPoint)
    End Function
End Class
