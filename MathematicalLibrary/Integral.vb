Imports MathematicalLibrary.Spline3

Public Class Integral

    Public Shared Function ИнтегрированиеРадиальнойЭпюры(ByVal InputX() As Double, ByVal inputData() As Double) As Double
        'на координатах равномерного массива InputX 
        'dtX вычисляется как первое минус нулевое значение массива
        'ширина основания равна последнему значению массива плюс dtX
        Dim dtX As Double = InputX(1) - InputX(0)
        Dim initialCondition As Double = inputData(0)
        Dim finalCondition As Double = inputData(UBound(inputData))
        Dim SimpsonRule As Double() = SignalProcessing.Integrate(inputData, dtX, initialCondition, finalCondition)
        Return SimpsonRule(UBound(SimpsonRule)) / (InputX(UBound(InputX)) + dtX)
    End Function

    Public Shared Function ИнтегрированиеРадиальнойЭпюрыНаПроизвольныхКоординатах(ByVal InputX() As Double, ByVal inputData() As Double, ByVal Ширина As Double) As Double
        Dim ЧислоТермопар As Integer = InputX.Count
        Dim НоваяРазмерность As Integer = ЧислоТермопар + 2

        Dim TblКоэффициенты As Double(,) = Nothing
        Dim arrX(ЧислоТермопар + 1), arrY(ЧислоТермопар + 1) As Double
        'скопировать входной массив (с 1 по предпоследний будут заполнены)
        Array.Copy(InputX, 0, arrX, 1, ЧислоТермопар)
        Array.Copy(inputData, 0, arrY, 1, ЧислоТермопар)
        arrX(ЧислоТермопар + 1) = Ширина 'начало  0, последний - ширина
        'найти значения температур на стенках

        'найдем таблицу коэффициентов один раз
        Spline3BuildTable(ЧислоТермопар, 2, InputX, inputData, 0, 0, TblКоэффициенты) 'UBound(InputX) + 1
        arrY(0) = Spline3Interpolate(ЧислоТермопар, TblКоэффициенты, 0) 'Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, 0)
        arrY(ЧислоТермопар + 1) = Spline3Interpolate(ЧислоТермопар, TblКоэффициенты, Ширина)

        'теперь есть 2 массива arrX и arrY
        'надо равномерно на основании сплайна найти значения на промежуточных координатах
        Dim newInputX(100), newInputY(100) As Double
        Dim dtX As Double = Ширина / 100
        'найдем таблицу коэффициентов уже расширенных массивов
        Spline3BuildTable(НоваяРазмерность, 2, arrX, arrY, 0, 0, TblКоэффициенты) 'Spline3BuildTable(UBound(arrX) + 1, 2, arrX, arrY, 0, 0, TblКоэффициенты)
        For I As Integer = 0 To 100 'CInt(Ширина / dtX)
            If I > 0 Then newInputX(I) = newInputX(I - 1) + dtX
            newInputY(I) = Spline3Interpolate(НоваяРазмерность, TblКоэффициенты, newInputX(I)) 'newInputY(I) = Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, newInputX(I))
        Next
        Return ИнтегрированиеРадиальнойЭпюры(newInputX, newInputY)
    End Function

    Public Shared Function ИнтегрированиеРадиальнойЭпюрыНаРавномерныхКоординатах(ByVal InputX() As Double, ByVal inputData() As Double, ByVal Ширина As Double) As Double
        '***********************************
        'расчет интегральной температура газа по мерному сечению на поясе
        'внимание этот расчет для равномерного положения термопар относительно друг друга и от стенок
        'т.е. все расстояния равномерны!!!
        '**********************************
        Dim ЧислоТермопар As Integer = InputX.Count

        Dim arrX(InputX.Length + 1), arrY(InputX.Length + 1) As Double
        'скопировать входной массив (с 1 по предпоследний будут заполнны)
        Array.Copy(InputX, 0, arrX, 1, ЧислоТермопар)
        Array.Copy(inputData, 0, arrY, 1, ЧислоТермопар)
        arrX(ЧислоТермопар + 1) = Ширина 'начало  0, последний - ширина

        'найти значения температур на стенках
        Dim Ystart As Double, Yend As Double
        НайтиЗначенияТемпературНаСтенкахInterpolate(InputX, inputData, Ширина, Ystart, Yend)
        arrY(0) = Ystart
        arrY(ЧислоТермопар + 1) = Yend

        Return ИнтегрированиеРадиальнойЭпюры(arrX, arrY)
    End Function

    Public Shared Function НайтиЗначенияТемпературыНаСтенке(ByVal inputX() As Double, ByVal inputY() As Double, ByVal координатаСтенки As Double) As Double
        'массивы начинаются с 0
        'найти значения температур на стенках
        Dim TblКоэффициенты As Double(,) = Nothing
        'найдем таблицу коэффициентов
        Spline3.Spline3BuildTable(UBound(inputX) + 1, 2, inputX, inputY, 0, 0, TblКоэффициенты)
        Return Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, координатаСтенки)
    End Function

    Public Shared Sub НайтиЗначенияТемпературНаСтенкахInterpolate(ByVal inputX() As Double, ByVal inputY() As Double, ByVal ширинаМерногоУчастка As Double, ByRef Ystart As Double, ByRef Yend As Double)
        'найти значения температур на стенках
        Dim TblКоэффициенты As Double(,) = Nothing

        'найдем таблицу коэффициентов
        Spline3.Spline3BuildTable(UBound(inputX) + 1, 2, inputX, inputY, 0, 0, TblКоэффициенты)

        'найти значения температур на стенках
        Ystart = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, 0)
        Yend = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, ширинаМерногоУчастка)
    End Sub

    Public Shared Sub НайтиЗначенияТемпературНаСтенкахЛинейно(ByVal inputX() As Double, ByVal inputY() As Double, ByVal ширинаМерногоУчастка As Double, ByRef Ystart As Double, ByRef Yend As Double)
        'массивы с 0 до 9
        'найти значения температур на стенках
        Ystart = inputY(0) + (inputY(1) - inputY(0)) * (0 - inputX(0)) / (inputX(1) - inputX(0))
        Yend = inputY(inputY.Length - 2) + (inputY(inputY.Length - 1) - inputY(inputY.Length - 2)) * (ширинаМерногоУчастка - inputX(inputX.Length - 2)) / (inputX(inputX.Length - 1) - inputX(inputX.Length - 2))
    End Sub
    '            y(0) = Ystart
    '        y(ЧислоТермопар + 1) = Yend

    '    y(0) = arrТемператураТ340(1) + (arrТемператураТ340(2) - arrТемператураТ340(1)) * (x(0) - arrТекущиеКоординаты(1)) / (arrТекущиеКоординаты(2) - arrТекущиеКоординаты(1))
    'y(11) = arrТемператураТ340(9) + (arrТемператураТ340(10) - arrТемператураТ340(9)) * (x(11) - arrТекущиеКоординаты(9)) / (arrТекущиеКоординаты(10) - arrТекущиеКоординаты(9))
End Class
