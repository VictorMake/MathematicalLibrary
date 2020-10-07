Public Class Spline3
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Бикубическое ресэмплирование (бикубическая интерполяция) 
    'Рассмотрим следующую задачу - у нас есть равномерная сетка размером M1 *N1  узлов. В узлах этой сетки нам известны значения некоторой функции f(x, y). Требуется перейти от сетки размером M1 *N1  к равномерной сетке размером M2 *N2  узлов, покрывающей ту же область (новая сетка может быть как более, так и менее плотная, чем старая), и вычислить значения функции на новой сетке. В качестве примера можно привести изменение размера растровой картинки (здесь функцией, определенной на сетке, являются цвета пикселей). 
    'Это задача интерполяции - по известным значениям в одних точках получить значения функции в других точках. Но данный случай является особым - нам требуется вычислять значение функции не в произвольных точках, а в узлах сетки, и вычисления проводятся только один раз, после чего мы получаем новую сетку и больше не возвращаемся к ресурсоемкой интерполяции. Поэтому для решения данной задачи имеет смысл разработать специализированную процедуру. 
    'Бикубическая интерполяция, реализованная в данном алгоритме, отличается от билинейной интерполяции тем, что используемый интерполирующий бикубический сплайн имеет гладкую первую производную и непрерывную вторую производную. 
    'Описание алгоритма 
    'Задачу двухмерной интерполяции бикубическим сплайном мы сводим к последовательности одномерных интерполяций кубическими сплайнами. Сначала изменяется горизонтальный размер исходной сетки, затем происходит изменение вертикального размера. На каждом шаге алгоритма нам требуется проводить только одномерную интерполяцию, с чем легко справляется используемый алгоритм интерполяции кубическими сплайнами, который также требуется скачать с сайта. 
    'Описание подпрограммы BicubicResample 
    'Входные параметры: 
    'OldWidth, OldHeight - старый размер сетки 
    'NewWidth, NewHeight - новый размер сетки 
    'A - массив значений функции на старой сетке. Нумерация элементов [0..OldHeight-1, 0..OldWidth-1] 
    'Выходные параметры: 
    'B - массив значений функции на новой сетке. Нумерация элементов [0..NewHeight-1, 0..NewWidth-1] 
    'Допустимые значения параметров: 
    'OldWidth > 1 
    'OldHeight > 1 
    'NewWidth > 1 
    'NewHeight > 1 
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Эти подпрограммы должны быть определены программистом
    ' Sub Spline3BuildTable(ByVal N As integer, _
    '         ByVal DiffN As integer, _
    '         ByRef x_() As Double, _
    '         ByRef y_() As Double, _
    '         ByVal BoundL As Double, _
    '         ByVal BoundR As Double, _
    '         ByRef ctbl() As Double)
    ' Function Spline3Interpolate(ByVal N As integer, _
    '         ByRef c() As Double, _
    '         ByVal X As Double) As Double


    'Подпрограммы
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Ресэмплирование бикубическим сплайном.
    '
    '    Процедура   получает   значения   функции    на     сетке
    'OldWidth*OldHeight и путем интерполяции бикубическим сплайном
    'вычисляет значения функции в узлах  сетки  размером NewWidth*
    'NewHeight. Новая  сетка  может  быть как более, так  и  менее
    'плотная, чем старая.
    '
    'Входные параметры:
    '    OldWidth    - старый размер сетки
    '    OldHeight   - старый размер сетки
    '    NewWidth    - новый размер сетки
    '    NewHeight   - новый размер сетки
    '    A           - массив значений функции на старой сетке.
    '                  Нумерация элементов [0..OldHeight-1,
    '                  0..OldWidth-1]
    '
    'Выходные параметры:
    '    B           - массив значений функции на новой сетке.
    '                  Нумерация элементов [0..NewHeight-1,
    '                  0..NewWidth-1]
    '
    'Допустимые значения параметров:
    '    OldWidth>1, OldHeight>1
    '    NewWidth>1, NewHeight>1
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Sub BicubicResample(ByVal OldWidth As Integer, _
             ByVal OldHeight As Integer, _
             ByVal NewWidth As Integer, _
             ByVal NewHeight As Integer, _
             ByRef A(,) As Double, _
             ByRef B(,) As Double)
        Dim Buf As Double(,)
        Dim Tbl As Double(,) = Nothing
        Dim X As Double()
        Dim Y As Double()
        Dim MW As Integer
        Dim MH As Integer
        Dim I As Integer
        Dim J As Integer

        'MW = MaxInt(OldWidth, NewWidth)
        MW = Math.Max(OldWidth, NewWidth)
        'MH = MaxInt(OldHeight, NewHeight)
        MH = Math.Max(OldHeight, NewHeight)

        'ReDim_B(NewHeight - 1, NewWidth - 1)
        'ReDim_Buf(MH - 1, MW - 1)
        Re.Dim(B, NewHeight - 1, NewWidth - 1)
        Re.Dim(Buf, MH - 1, MW - 1)
        'ReDim_X(MaxInt(MW, MH) - 1)
        'ReDim_Y(MaxInt(MW, MH) - 1)
        'ReDim_X(Math.Max(MW, MH) - 1)
        'ReDim_Y(Math.Max(MW, MH) - 1)
        Re.Dim(X, Math.Max(MW, MH) - 1)
        Re.Dim(Y, Math.Max(MW, MH) - 1)

        For I = 0 To OldHeight - 1 Step 1
            For J = 0 To OldWidth - 1 Step 1
                X(J) = J / (OldWidth - 1)
                Y(J) = A(I, J)
            Next J

            Spline3BuildTable(OldWidth, 2, X, Y, 0, 0, Tbl)

            For J = 0 To NewWidth - 1 Step 1
                Buf(I, J) = Spline3Interpolate(OldWidth, Tbl, J / (NewWidth - 1))
            Next J
        Next I

        For J = 0 To NewWidth - 1 Step 1
            For I = 0 To OldHeight - 1 Step 1
                X(I) = I / (OldHeight - 1)
                Y(I) = Buf(I, J)
            Next I

            Spline3BuildTable(OldHeight, 2, X, Y, 0, 0, Tbl)

            For I = 0 To NewHeight - 1 Step 1
                B(I, J) = Spline3Interpolate(OldHeight, Tbl, I / (NewHeight - 1))
            Next I
        Next J
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Интерполяция функции кубическими сплайнами 
    'Интерполяция кубическими сплайнами - это быстрый, эффективный и устойчивый способ интерполяции функций, который является основным конкурентом полиномиальной интерполяции. В его основе лежит следующая идея - интервал интерполяции разбивается на небольшие отрезки, на каждом из которых функция задается полиномом третьей степени. Коэффициенты полинома подбираются так, что на границах интервалов обеспечивается непрерывность функции, её первой и второй производных.
    ' Также есть возможность задать граничные условия - значения первой или второй производной на границах интервала. Если значения одной из производных на границе известны, то задав их, мы получаем крайне точную интерполяционную схему. Если значения неизвестны, то можно задать вторую производную на границе равной нолю и получить достаточно хорошие результаты. 
    'Теперь о математической части. Пусть заданы точки x1 , x2 , ..., xn  и соответствующие им значения y1 , y2 , ..., yn  функции f(x). На каждом из отрезков [xi , xi+1 ], i=1, 2, ..., n-1 функцию приближаем при помощи полинома третьей степени: 
    'S(x) = yi  + c1i (x-xi ) + c2i (x-xi ) 2 + c3i (x-xi ) 3, xi  < x < xi+1 
    'Для вычисления коэффициентов c1i , c2i , c3i , i = 1, 2, ..., n-1 решается система линейных уравнений, построенная из условия непрерывности производной S'(x) в узлах сетки и дополнительных краевых условий на вторую производную, которые имеют вид: 
    '2*S''1  + b1 *S''2  = b2 
    'b3 *S''N-1  + 2*S''N  = b4 
    'здесь возможны два случая. Случай первый, когда известны значения первой производной в краевых точках (y'1  = y'(x1 ), y'n  = y'(xn )), тогда следует положить: 
    'b1  = 1, b2  = (6/(x2 -x1 )) * ((y2 -y1 ) / (x2 -x1 )-y'1 ),
    'b3  = 1, b4  = (6/(xn -xn-1 )) * (y'n  - (yn -yn-1 )/(xn -xn-1 ))
    'Случай второй, когда известны значения второй производной (y''1  = y''(x1 ), y''n  = y''(xn )), тогда полагаем: 
    'b1  = 0, b2  = 2*y''1 
    'b3  = 0, b4  = 2*y''N 
    'Описание подпрограммы Spline3BuildTable
    'Процедура Spline3BuildTable служит для постройки таблицы коэффициентов кубического сплайна по заданным точкам и граничным условиям, накладываемым на производные. В дальнейшем постренная таблица используется функцией Spline3Interpolate. 
    'Входные параметры: 

    'N - число точек 
    'DiffN - тип граничного условия. 1 соответствует граничным условиям накладываемым на первые производные, 2 - на вторые. 
    'xs - массив абсцисс опорных точек с номерами от 0 до N-1. Точки могут быть неупорядочены по возрастанию, алгоритм отсортирует их перед расчетами. 
    'ys - массив ординат опорных точек с номерами от 0 до N-1. 
    'BoundL - левое граничное условие. Если DiffN равно 1, то первая производная на левой границе равна BoundL, иначе BoundL равна вторая производная. 
    'BoundR - аналогично BoundL 
    'Выходные значения: 

    'ctbl - в этот массив помещается таблица коэффициентов сплайна. Эта таблица используется функцией Spline3Interpolate. 
    'Описание подпрограммы Spline3Interpolate
    'Подпрограмма вычисляет значение сплайна в данной точке на основе таблицы коэффициентов, расчитанной в подпрограмме Spline3BuildTable. 

    'Входные параметры: 

    'N - число точек, переданных в процедуру Spline3BuildTable 
    'C - таблица коэффициентов, построенная процедурой Spline3BuildTable 
    'X - точка, в которой ведем расчет 
    'Результат - значение кубического сплайна, заданного таблицей C, в точке X. 

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Подпрограммы
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Процедура Spline3BuildTable служит  для  постройки  таблицы  коэффициентов
    'кубического сплайна по заданным точкам и граничным условиям, накладываемым
    'на производные.
    '
    'В дальнейшем постренная таблица используется функцией Spline3Interpolate.
    '
    'Параметры:
    '    N       - число точек
    '    DiffN   - тип граничного условия. "1" соответствует граничным условиям
    '              накладываемым на первые производные, "2" - на вторые.
    '    xs      - массив абсцисс опорных точек с номерами от 0 до N-1
    '    ys      - массив ординат опорных точек с номерами от 0 до N-1
    '    BoundL  - левое граничное условие. Если DiffN равно 1, то первая произ-
    '              водная  на  левой  границе  равна BoundL, иначе BoundL равна
    '              вторая.
    '    BoundR  - аналогично BoundL
    '
    '
    'Выходные значения:
    '    ctbl    - в этот массив помещается таблица коэффициентов сплайна.
    '              Эта таблица используется функцией Spline3Interpolate.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Sub Spline3BuildTable(ByVal N As Integer,
             ByVal DiffN As Integer,
             ByRef x_() As Double,
             ByRef y_() As Double,
             ByVal BoundL As Double,
             ByVal BoundR As Double,
             ByRef ctbl(,) As Double)

        Dim X As Double()
        Dim y As Double()
        Dim c As Boolean
        'Dim E As Integer
        Dim G As Integer
        Dim Tmp As Double
        Dim nxm1 As Integer
        Dim I As Integer
        Dim J As Integer
        Dim DX As Double
        Dim DXJ As Double
        Dim DYJ As Double
        Dim DXJP1 As Double
        Dim DYJP1 As Double
        Dim DXP As Double
        'Dim DYP As Double
        Dim YPPA As Double
        Dim YPPB As Double
        Dim PJ As Double
        Dim b1 As Double
        Dim b2 As Double
        Dim b3 As Double
        Dim b4 As Double
        X = x_
        y = y_

        N -= 1
        G = (N + 1) \ 2

        Do
            I = G
            Do
                J = I - G
                c = True
                Do
                    If X(J) <= X(J + G) Then
                        c = False
                    Else
                        Tmp = X(J)
                        X(J) = X(J + G)
                        X(J + G) = Tmp
                        Tmp = y(J)
                        y(J) = y(J + G)
                        y(J + G) = Tmp
                    End If
                    J -= 1
                Loop Until Not (J >= 0 AndAlso c)
                I += 1
            Loop Until Not I <= N
            G \= 2
        Loop Until Not G > 0

        'ReDim_ctbl(4, N)
        Re.Dim(ctbl, 4, N)
        N += 1

        If DiffN = 1 Then
            b1 = 1
            b2 = 6 / (X(1) - X(0)) * ((y(1) - y(0)) / (X(1) - X(0)) - BoundL)
            b3 = 1
            b4 = 6 / (X(N - 1) - X(N - 2)) * (BoundR - (y(N - 1) - y(N - 2)) / (X(N - 1) - X(N - 2)))
        Else
            b1 = 0
            b2 = 2 * BoundL
            b3 = 0
            b4 = 2 * BoundR
        End If

        nxm1 = N - 1

        If N >= 2 Then
            If N > 2 Then
                DXJ = X(1) - X(0)
                DYJ = y(1) - y(0)
                J = 2
                Do While J <= nxm1
                    DXJP1 = X(J) - X(J - 1)
                    DYJP1 = y(J) - y(J - 1)
                    DXP = DXJ + DXJP1
                    ctbl(1, J - 1) = DXJP1 / DXP
                    ctbl(2, J - 1) = 1 - ctbl(1, J - 1)
                    ctbl(3, J - 1) = 6 * (DYJP1 / DXJP1 - DYJ / DXJ) / DXP
                    'If Double.IsInfinity(ctbl(3, J - 1)) Then Stop
                    DXJ = DXJP1
                    DYJ = DYJP1
                    J += 1
                Loop
            End If

            ctbl(1, 0) = -(b1 / 2)
            ctbl(2, 0) = b2 / 2

            If N <> 2 Then
                J = 2
                Do While J <= nxm1
                    PJ = ctbl(2, J - 1) * ctbl(1, J - 2) + 2
                    ctbl(1, J - 1) = -(ctbl(1, J - 1) / PJ)
                    ctbl(2, J - 1) = (ctbl(3, J - 1) - ctbl(2, J - 1) * ctbl(2, J - 2)) / PJ
                    'If Double.IsNaN(ctbl(2, J - 1)) Then Stop
                    J += 1
                Loop
            End If

            YPPB = (b4 - b3 * ctbl(2, nxm1 - 1)) / (b3 * ctbl(1, nxm1 - 1) + 2)
            I = 1

            'If Double.IsNaN(YPPB) Then Stop
            Do While I <= nxm1
                J = N - I
                YPPA = ctbl(1, J - 1) * YPPB + ctbl(2, J - 1)
                DX = X(J) - X(J - 1)
                ctbl(3, J - 1) = (YPPB - YPPA) / DX / 6
                ctbl(2, J - 1) = YPPA / 2
                ctbl(1, J - 1) = (y(J) - y(J - 1)) / DX - (ctbl(2, J - 1) + ctbl(3, J - 1) * DX) * DX
                'If Double.IsNaN(ctbl(1, J - 1)) Then Stop
                YPPB = YPPA
                I += 1
            Loop

            For I = 1 To N Step 1
                ctbl(0, I - 1) = y(I - 1)
                ctbl(4, I - 1) = X(I - 1)
            Next I
        End If
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Вычисление значения сплайна в точке по таблице коэффициентов
    '
    'Функция Spline3Interpolate по построенной процедурой Spline3BuildTable
    'таблице коэффициентов вычисляет значение интерполируемой функции в
    'заданной точке.
    '
    'Параметры:
    '    N       - число точек, переданных в процедуру Spline3BuildTable
    '    C       - таблица коэффициентов, построенная процедурой
    '              Spline3BuildTable
    '    X       - точка, в которой ведем расчет
    '
    'Результат:
    '    значение кубического сплайна, заданного таблицей C, в точке X
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Function Spline3Interpolate(ByVal N As Integer,
             ByRef c(,) As Double,
             ByVal X As Double) As Double

        Dim I As Integer
        Dim L As Integer
        Dim Half As Integer
        Dim First As Integer
        Dim Middle As Integer

        N -= 1
        L = N
        First = 0

        Do While L > 0
            Half = L \ 2
            Middle = First + Half
            If c(4, Middle) < X Then
                First = Middle + 1
                L = L - Half - 1
            Else
                L = Half
            End If
        Loop

        I = First - 1

        If I < 0 Then
            I = 0
        End If

        Return c(0, I) + (X - c(4, I)) * (c(1, I) + (X - c(4, I)) * (c(2, I) + c(3, I) * (X - c(4, I))))
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Билинейное ресэмплирование (билинейная интерполяция) 
    'Рассмотрим следующую задачу - у нас есть равномерная сетка размером M1 *N1  узлов. В узлах этой сетки нам известны значения некоторой функции f(x, y). Требуется перейти от сетки размером M1 *N1  к равномерной сетке размером M2 *N2  узлов, покрывающей ту же область (новая сетка может быть как более, так и менее плотная, чем старая), и вычислить значения функции на новой сетке. В качестве примера можно привести изменение размера растровой картинки (здесь функцией, определенной на сетке, являются цвета пикселей). 
    'Это задача интерполяции - по известным значениям в одних точках получить значения функции в других точках. Но данный случай является особым - нам требуется вычислять значение функции не в произвольных точках, а в узлах сетки, и вычисления проводятся только один раз, после чего мы получаем новую сетку и больше не возвращаемся к ресурсоемкой интерполяции. Поэтому для решения данной задачи имеет смысл разработать специализированную процедуру. 
    'Во многих случаях оказывается достаточно той точности, которую обеспечивает билинейная интерполяция. Скажем, для изменения размера изображений этого хватает. Тем не менее, следует отметить, что используемая в данном способе кусочно-билинейная интерполирующая функция имеет разрывный градиент, что оказывается недопустимым в ряде случаев. В задачах, требующих более гладкой интерполяции, лучше использовать другие методы, например бикубическое ресэмплирование. 
    'Описание алгоритма 
    'Сам алгоритм крайне прост - мы обходим точки новой сетки. Для каждой точки смотрим, в какой квадрат старой сетки она попадает. Пусть точка попала в квадрат, образованный узлами (xL , yC ), (xL+1 , yC ), (xL , yC+1 ), (xL+1 , yC+1 ). Применяем к координатам точки (x, y) преобразование u = (x-xL )/(xL+1 -xL ), t = (y-yC )/(yC+1 -yC ). Полученные координаты лежат в диапазоне [0, 1]. Значение функции в новом узле расчитывается, как fnew  = (1-t)(1-u)*fL,C +t(1-u)*fL,C+1 +tu*fL+1,C+1 +(1-t)u*fL+1,C . 
    'Описание подпрограммы BilinearResample 
    'Входные параметры: 
    'OldWidth, OldHeight - старый размер сетки 
    'NewWidth, NewHeight - новый размер сетки 
    'A - массив значений функции на старой сетке. Нумерация элементов [0..OldHeight-1, 0..OldWidth-1] 
    'Выходные параметры: 
    'B - массив значений функции на новой сетке. Нумерация элементов [0..NewHeight-1, 0..NewWidth-1] 
    'Допустимые значения параметров: 
    'OldWidth > 1 
    'OldHeight > 1 
    'NewWidth > 1 
    'NewHeight > 1 
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'Подпрограммы
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Билинейное ресэмплирование.
    '
    '    Процедура   получает   значения   функции    на     сетке
    'OldWidth*OldHeight и путем билинейной интерполяции  вычисляет
    'значения функции в узлах  сетки  размером NewWidth*NewHeight.
    'Новая  сетка  может  быть как более, так и менее плотная, чем
    'старая.
    '
    'Входные параметры:
    '    OldWidth    - старый размер сетки
    '    OldHeight   - старый размер сетки
    '    NewWidth    - новый размер сетки
    '    NewHeight   - новый размер сетки
    '    A           - массив значений функции на старой сетке.
    '                  Нумерация элементов [0..OldHeight-1,
    '                  0..OldWidth-1]
    '
    'Выходные параметры:
    '    B           - массив значений функции на новой сетке.
    '                  Нумерация элементов [0..NewHeight-1,
    '                  0..NewWidth-1]
    '
    'Допустимые значения параметров:
    '    OldWidth>1, OldHeight>1
    '    NewWidth>1, NewHeight>1
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Shared Sub BilinearResample(ByVal OldWidth As Integer,
             ByVal OldHeight As Integer,
             ByVal NewWidth As Integer,
             ByVal NewHeight As Integer,
             ByRef A(,) As Double,
             ByRef B(,) As Double)

        Dim I As Integer
        Dim J As Integer
        Dim L As Integer
        Dim C As Integer
        Dim T As Double
        Dim U As Double
        Dim Tmp As Double

        'ReDim_B(NewHeight - 1, NewWidth - 1)
        Re.Dim(B, NewHeight - 1, NewWidth - 1)

        For I = 0 To NewHeight - 1 Step 1
            For J = 0 To NewWidth - 1 Step 1
                Tmp = I / (NewHeight - 1) * (OldHeight - 1)
                L = CInt(Int(Tmp))

                If L < 0 Then
                    L = 0
                Else
                    If L >= OldHeight - 1 Then
                        L = OldHeight - 2
                    End If
                End If

                U = Tmp - L
                Tmp = J / (NewWidth - 1) * (OldWidth - 1)
                C = CInt(Int(Tmp))

                If C < 0 Then
                    C = 0
                Else
                    If C >= OldWidth - 1 Then
                        C = OldWidth - 2
                    End If
                End If

                T = Tmp - C
                B(I, J) = (1 - T) * (1 - U) * A(L, C) + T * (1 - U) * A(L, C + 1) + T * U * A(L + 1, C + 1) + (1 - T) * U * A(L + 1, C)
            Next J
        Next I
    End Sub


    '    Константы
    'MachineEpsilon
    'Эта константа определяет точность машинных операций, т.е. минимальное число, такое, что 1+MachineEpsilon≠1 на данной разрядной сетке. Константа может быть взята "с запасом", т.е. реальная точность может быть ещё выше, но по поводу указанного числа можно быть уверенным, что всё будет ОК. 

    'MaxRealNumber
    'Эта константа определяет максимальное положительное вещественное число, представимое на данной машине. Константа может быть взята "с запасом", т.е. реальная граница может быть ещё выше, но по поводу указанного числа можно быть уверенным, что всё будет ОК. 

    'MinRealNumber
    'Эта константа определяет минимальное положительное вещественное число, представимое на данной машине. Константа может быть взята "с запасом", т.е. реальная граница может быть ещё ниже, но по поводу указанного числа можно быть уверенным, что всё будет ОК. 

    'Функции
    '    Public Function MaxReal(ByVal M1 As Double, ByVal M2 As Double) As Double
    'Возвращает максимум из двух вещественных чисел. 

    'Public Function MinReal(ByVal M1 As Double, ByVal M2 As Double) As Double
    'Возвращает минимум из двух вещественных чисел. 

    'Public Function MaxInt(ByVal M1 As Long, ByVal M2 As Long) As Long
    'Возвращает максимум из двух целых чисел. 

    'Public Function MinInt(ByVal M1 As Long, ByVal M2 As Long) As Long
    'Возвращает минимум из двух целых чисел. 

    'Public Function ArcSin(ByVal X As Double) As Double
    'Арксинус (возвращает угол в радианах). 

    'Public Function ArcCos(ByVal X As Double) As Double
    'Арккосинус (возвращает угол в радианах). 

    'Public Function SinH(ByVal X As Double) As Double
    'Гиперболический синус. 

    'Public Function CosH(ByVal X As Double) As Double
    'Гиперболический косинус. 

    'Public Function TanH(ByVal X As Double) As Double
    'Гиперболический тангенс. 

    'Public Function Pi() As Double
    'Возвращает значение числа π. 

    'Public Function Power(ByVal Base As Double, ByVal Exponent As Double) As Double
    'Возвращает Base в степени Exponent (введено для совместимости). 

    'Public Function Square(ByVal X As Double) As Double
    'Возвращает x2. 

    'Public Function Log10(ByVal X As Double) As Double
    'Возвращает десятичный логарифм X. 

    'Public Function Ceil(ByVal X As Double) As Double
    'Самое маленькое целое число, большее или равное X. 

    'Public Function RandomInteger(ByVal X As Long) As Long
    'Возвращает случайное целое число в полуинтервале [0, I). 

    'Public Function Atn2(ByVal Y As Double, ByVal X As Double) As Double
    'Аргумент комплексного числа X + iY. В диапазоне от -π до π. 

    'Public Const MachineEpsilon = 0.0000000000000005
    'Public Const MaxRealNumber = 1.0E+300
    'Public Const MinRealNumber = 1.0E-300

    'Private Const BigNumber As Double = 1.0E+70
    'Private Const SmallNumber As Double = 1.0E-70
    'Private Const PiNumber As Double = 3.14159265358979

    'Public Function MaxReal(ByVal M1 As Double, ByVal M2 As Double) As Double
    '    If M1 > M2 Then
    '        MaxReal = M1
    '    Else
    '        MaxReal = M2
    '    End If
    'End Function

    'Public Function MinReal(ByVal M1 As Double, ByVal M2 As Double) As Double
    '    If M1 < M2 Then
    '        MinReal = M1
    '    Else
    '        MinReal = M2
    '    End If
    'End Function

    'Private Shared Function MaxInt(ByVal M1 As Integer, ByVal M2 As Integer) As Integer
    '    If M1 > M2 Then
    '        MaxInt = M1
    '    Else
    '        MaxInt = M2
    '    End If
    'End Function

    'Private Shared Function MinInt(ByVal M1 As Integer, ByVal M2 As Integer) As Integer
    '    If M1 < M2 Then
    '        MinInt = M1
    '    Else
    '        MinInt = M2
    '    End If
    'End Function

    'Public Function ArcSin(ByVal X As Double) As Double
    '    Dim T As Double
    '    T = Sqr(1 - X * X)
    '    If T < SmallNumber Then
    '        ArcSin = Atn(BigNumber * Sgn(X))
    '    Else
    '        ArcSin = Atn(X / T)
    '    End If
    'End Function

    'Public Function ArcCos(ByVal X As Double) As Double
    '    Dim T As Double
    '    T = Sqr(1 - X * X)
    '    If T < SmallNumber Then
    '        ArcCos = Atn(BigNumber * Sgn(-X)) + 2 * Atn(1)
    '    Else
    '        ArcCos = Atn(-X / T) + 2 * Atn(1)
    '    End If
    'End Function

    'Public Function SinH(ByVal X As Double) As Double
    '    SinH = (Exp(X) - Exp(-X)) / 2
    'End Function

    'Public Function CosH(ByVal X As Double) As Double
    '    CosH = (Exp(X) + Exp(-X)) / 2
    'End Function

    'Public Function TanH(ByVal X As Double) As Double
    '    TanH = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
    'End Function

    'Public Function Pi() As Double
    '    Pi = PiNumber
    'End Function

    'Public Function Power(ByVal Base As Double, ByVal Exponent As Double) As Double
    '    Power = Base ^ Exponent
    'End Function

    'Public Function Square(ByVal X As Double) As Double
    '    Square = X * X
    'End Function

    'Public Function Log10(ByVal X As Double) As Double
    '    Log10 = Log(X) / Log(10)
    'End Function

    'Public Function Ceil(ByVal X As Double) As Double
    '    Ceil = -Int(-X)
    'End Function

    'Public Function RandomInteger(ByVal X As Integer) As Integer
    '    RandomInteger = Int(Rnd() * X)
    'End Function

    'Public Function Atn2(ByVal Y As Double, ByVal X As Double) As Double
    '    If SmallNumber * Abs(Y) < Abs(X) Then
    '        If X < 0 Then
    '            If Y = 0 Then
    '                Atn2 = Pi()
    '            Else
    '                Atn2 = Atn(Y / X) + Pi() * Sgn(Y)
    '            End If
    '        Else
    '            Atn2 = Atn(Y / X)
    '        End If
    '    Else
    '        Atn2 = Sgn(Y) * Pi() / 2
    '    End If
    'End Function

    Public Shared Function Аппроксимация(ByVal Nkon As Integer, ByRef Xg() As Double, ByRef Yg() As Double, ByRef M() As Double, ByVal Xtek As Double) As Double
        '    ' Функция расчитывает значение Y по текущему Xtek
        '    ' Nkon последнее значение X, Xg(1 to Nkon) значения по X, Yg(1 to Nkon) результат функции
        '    ' M(1 to Nkon) коэффициенты по результату работы процедуры ВводАппроксимация
        Dim I, J As Integer
        Dim P, f, D, H, R As Double

        If Xtek > Xg(Nkon) Then
            D = Xg(Nkon) - Xg(Nkon - 1)
            f = D * M(Nkon - 1) / 6 + (Yg(Nkon) - Yg(Nkon - 1)) / D
            Return f * (Xtek - Xg(Nkon)) + Yg(Nkon)
            Exit Function
        End If

        If Xtek <= Xg(1) Then
            D = Xg(2) - Xg(1)
            f = -D * M(2) / 6 + (Yg(2) - Yg(1)) / D
            Return f * (Xtek - Xg(1)) + Yg(1)
            Exit Function
        End If

        Do
            I += 1
        Loop While Xtek > Xg(I)

        J = I - 1 : D = Xg(I) - Xg(J) : H = Xtek - Xg(J) : R = Xg(I) - Xtek
        P = D * D / 6 : f = (M(J) * R * R * R + M(I) * H * H * H) / 6 / D

        Return f + ((Yg(J) - M(J) * P) * R + (Yg(I) - M(I) * P) * H) / D
    End Function

    Public Shared Sub ВводАппроксимация(ByVal Nkon As Integer, ByRef Xg() As Double, ByRef Yg() As Double, ByRef M() As Double)
        '    ' Процедура по входным массивам Xg(1 to Nkon) значения по X, Yg(1 to Nkon) результат функции
        '    ' вычисляет коэффициенты M(1 to Nkon) для сплайн аппроксимации - функции Аппроксимация
        '    ' и заносит их в временный массив M() который можно переписать в общий массив для коэффициентов
        Dim L(Nkon) As Double
        Dim R(Nkon) As Double
        Dim S(Nkon) As Double
        Dim K As Integer 'I,
        Dim f, E, D, H, P As Double

        D = Xg(2) - Xg(1)
        E = (Yg(2) - Yg(1)) / D

        For K = 2 To Nkon - 1
            H = D
            D = Xg(K + 1) - Xg(K)
            f = E
            E = (Yg(K + 1) - Yg(K)) / D
            L(K) = D / (D + H)
            R(K) = 1 - L(K)
            S(K) = 6 * (E - f) / (H + D)
        Next K

        For K = 2 To Nkon - 1
            P = 1 / (R(K) * L(K - 1) + 2)
            L(K) = -L(K) * P
            S(K) = (S(K) - R(K) * S(K - 1)) * P
        Next K

        M(Nkon) = 0
        L(Nkon - 1) = S(Nkon - 1)
        M(Nkon - 1) = L(Nkon - 1)

        For K = Nkon - 2 To 1 Step -1
            L(K) = L(K) * L(K + 1) + S(K)
            M(K) = L(K)
        Next K
    End Sub

    Public Shared Function InterpLine(ByVal X() As Double, ByVal Y() As Double, ByVal XVal As Double) As Double
        Dim XINFAC, Y1, Y2, FVAL As Double
        Dim I, J, QPINDEX, N As Integer
        ' Linear interpolation subroutine
        ' y = f(x)

        ' Input
        '  N    = number of X and Y data points ( N >= 2 )
        '  X()  = vector of X data ( N rows )
        '  Y()  = vector of Y data ( N rows )
        '  XVAL = X argument

        ' Output
        '  FVAL = interpolated function value at XVAL
        N = UBound(X)
        'Внимание массив начинается с 1
        If XVal < X(1) Then
            FVAL = График(XVal, X(1), Y(1), X(2), Y(2))
        ElseIf XVal > X(N) Then
            FVAL = График(XVal, X(N - 1), Y(N - 1), X(N), Y(N))
        Else
            ' compute index and interpolation factor
            For I = 1 To N
                If (XVal <= X(I)) Then
                    If (I = 1) Then
                        XINFAC = 0.0
                        QPINDEX = 1
                    Else
                        J = I - 1
                        XINFAC = (XVal - X(J)) / (X(I) - X(J))
                        QPINDEX = J
                    End If
                    Exit For
                End If
            Next I

            Y1 = Y(QPINDEX)
            Y2 = Y(QPINDEX + 1)

            FVAL = Y1 + XINFAC * (Y2 - Y1)
        End If

        Return FVAL
    End Function

    Public Shared Function График(ByVal ЗаданX As Double, ByVal dblX1 As Double, ByVal dblY1 As Double, ByVal dblX2 As Double, ByVal dblY2 As Double) As Double
        If dblX2 - dblX1 = 0 Then
            График = dblY1
        Else
            График = dblY1 + (dblY2 - dblY1) * (ЗаданX - dblX1) / (dblX2 - dblX1)
        End If
    End Function
End Class
