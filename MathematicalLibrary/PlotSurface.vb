Public Class PlotSurface
    'Private xInit, dtX, xFinal As Double

    Public Shared Sub PlotSurfaceНаПроизвольных(ByVal длина As Integer, ByVal ширина As Integer, ByVal arrПоле(,) As Double, ByVal tArray() As Double)
        ''(ByVal InputArray As Double(,), ByVal XIstart As Double, ByVal XIfinish As Double, ByVal YJstart As Double, ByVal YJfinish As Double)

        ''Dim tArray() As Double = ПараметрыПоляНакопленные(conArrПоложениеГребенки).ВсеЗамеры
        ''Dim InputX(ПараметрыПоляНакопленные(conArrПоложениеГребенки).ВсеЗамеры.Length + 1) As Double
        'Dim InputX(tArray.Length + 1) As Double
        'Dim InputY(UBound(InputX)) As Double


        ''ЧислоПромежутков = от 0 до 16
        'Dim ВходнойМассив(54, 240) As Double
        ''Dim XIfinish As Integer = 54 ' / conШагСмещения 'ПараметрыПоляНакопленные.ЧислоПромежутков 2 1
        ''Dim YJfinish As Integer = 6
        'Dim J As Integer 'I, 


        'Dim TblКоэффициенты(,) As Double

        'Dim TempArray(54, 6) As Double


        ''Dim InputY() As Double '(UBound(InputX))



        ''делаем исскуственно н крайних значениях температуры как средние по полю T_интегрИтог
        'InputX(0) = 0
        'InputX(UBound(InputX)) = 54
        ''If blnИзмерениеПоТемпературам Then
        'InputY(0) = T_интегрИтог
        'InputY(UBound(InputY)) = T_интегрИтог
        ''Else
        ''    InputY(0) = dblИнтегральнаяПоПолю 'ПотериПолнДавленияИтегрИтог нельзя - это разность
        ''    InputY(UBound(InputY)) = dblИнтегральнаяПоПолю 'ПотериПолнДавленияИтегрИтог
        ''End If

        'For I As Integer = 0 To tArray.Length - 1
        '    InputX(I + 1) = tArray(I)
        'Next

        'For J = 0 To 6
        '    If J = 0 Then
        '        'InputY = ПараметрыПоляНакопленные(conArrЗначНаСтенка1).ВсеЗамеры
        '        'Array.Copy(ПараметрыПоляНакопленные(conArrЗначНаСтенка1).ВсеЗамеры, 1, InputY, 1, ПараметрыПоляНакопленные(conArrЗначНаСтенка1).ВсеЗамеры.Length)
        '        tArray = ПараметрыПоляНакопленные(conArrЗначНаСтенка1).ВсеЗамеры
        '        For I As Integer = 0 To tArray.Length - 1
        '            InputY(I + 1) = tArray(I)
        '        Next
        '    ElseIf J = 6 Then
        '        'InputY = ПараметрыПоляНакопленные(conArrЗначНаСтенка2).ВсеЗамеры
        '        'Array.Copy(ПараметрыПоляНакопленные(conArrЗначНаСтенка2).ВсеЗамеры, 1, InputY, 1, ПараметрыПоляНакопленные(conArrЗначНаСтенка2).ВсеЗамеры.Length)
        '        tArray = ПараметрыПоляНакопленные(conArrЗначНаСтенка2).ВсеЗамеры
        '        For I As Integer = 0 To tArray.Length - 1
        '            InputY(I + 1) = tArray(I)
        '        Next
        '    Else
        '        'Array.Copy(ПараметрыПоляНакопленные("arrПояс" & J.ToString).ВсеЗамеры, 1, InputY, 1, ПараметрыПоляНакопленные("arrПояс" & J.ToString).ВсеЗамеры.Length)
        '        'InputY = ПараметрыПоляНакопленные("arrПояс" & J.ToString).ВсеЗамеры
        '        tArray = ПараметрыПоляНакопленные("arrПояс" & J.ToString).ВсеЗамеры
        '        For I As Integer = 0 To tArray.Length - 1
        '            InputY(I + 1) = tArray(I)
        '        Next
        '    End If


        '    'найдем таблицу коэффициентов
        '    Spline3.Spline3BuildTable(UBound(InputX) + 1, 2, InputX, InputY, 0, 0, TblКоэффициенты)
        '    'через 1 мм проходим по каждому поясу
        '    'ПромКоорХ = 0
        '    'I = 0
        '    'Do While ПромКоорХ <= 54
        '    '    InputArray(I, J) = Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, ПромКоорХ)
        '    '    I += 1
        '    '    ПромКоорХ += conШагСмещения
        '    'Loop
        '    For Высота As Integer = 0 To 54
        '        TempArray(Высота, J) = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, Высота)
        '    Next
        'Next

        'InputX = New Double() {0, 40, 80, 120, 160, 200, 240}
        'ReDim_InputY(6)

        'For Высота As Integer = 0 To 54
        '    For J = 0 To 6
        '        InputY(J) = TempArray(Высота, J)
        '    Next

        '    'найдем таблицу коэффициентов
        '    Spline3.Spline3BuildTable(UBound(InputX) + 1, 2, InputX, InputY, 0, 0, TblКоэффициенты)
        '    'через 1 мм проходим по каждой высоте
        '    For Ширина As Integer = 0 To 240
        '        ВходнойМассив(Высота, Ширина) = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, Ширина)
        '    Next
        'Next


        'Dim x(54) As Double
        'Dim y(240) As Double
        'For I As Integer = 0 To 54
        '    x(I) = I
        'Next I
        'For I As Integer = 0 To 240
        '    y(I) = I
        'Next I

        ' ''растянуть и отдать на график
        ''Dim ПолеДавлений As Bicubic
        ''Dim ВыходнойМассив(,) As Double
        ''ПолеДавлений = New Bicubic(ВходнойМассив, 0, 54, 0, 240)
        ''ВыходнойМассив = ПолеДавлений.OutputArray


        'AxCWGraph3D1.Axes.Item("ZAxis").Minimum = Tgmin * 0.9
        'AxCWGraph3D1.Axes.Item("ZAxis").Maximum = Tgmax * 1.1
        'ComboPlotsList.SelectedIndex = 1
        'currentPlot = AxCWGraph3D1.Plots.Item(2)
        'AxCWGraph3D2.Plots.Item(2).Plot3DSurface(x, y, ВходнойМассив)
        'currentPlot.Plot3DSurface(x, y, ВходнойМассив)

        ''AxCWGraph3D1.Plot3DSimpleSurface(ВыходнойМассив)

    End Sub


    Public Shared Function ДостроитьПолеДавленийНаПромежуточных(ByVal Пояс As Integer, ByVal СтатическоеДавление As Double) As Double()
        'неизвестны значения на координатах 80, 120, 160, 200
        'с помощью аппроксимации попытаемся из вычислить
        Dim ВыходнойМассив(7) As Double
        'Dim InputX() As Double
        'Dim InputY() As Double
        'Dim TblКоэффициенты(,) As Double
        'Dim ЗначНаПромКоор, ПромКоорХ As Double
        ''проделать 5 раз
        ''dР310-1-1   dР310-3-5
        ''ВсеПараметрыПодсчета("dР310-1-" & Пояс.ToString).dblНакопленноеЗначение
        'InputX = New Double() {0, 40, 140, 240, 280}
        'InputY = New Double() {СтатическоеДавление, ВсеПараметрыПодсчета("dР310-1-" & Пояс.ToString).dblНакопленноеЗначение, ВсеПараметрыПодсчета("dР310-2-" & Пояс.ToString).dblНакопленноеЗначение, ВсеПараметрыПодсчета("dР310-3-" & Пояс.ToString).dblНакопленноеЗначение, СтатическоеДавление}

        ''найдем таблицу коэффициентов
        'Spline3.Spline3BuildTable(UBound(InputX) + 1, 2, InputX, InputY, 0, 0, TblКоэффициенты)

        ''на координате 40
        'ВыходнойМассив(1) = ВсеПараметрыПодсчета("dР310-1-" & Пояс.ToString).dblНакопленноеЗначение
        'ПромКоорХ = 80
        'ЗначНаПромКоор = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, ПромКоорХ)
        'ВыходнойМассив(2) = ЗначНаПромКоор
        'ПромКоорХ = 120
        'ЗначНаПромКоор = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, ПромКоорХ)
        'ВыходнойМассив(3) = ЗначНаПромКоор
        'ПромКоорХ = 160
        'ЗначНаПромКоор = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, ПромКоорХ)
        'ВыходнойМассив(4) = ЗначНаПромКоор
        'ПромКоорХ = 200
        'ЗначНаПромКоор = Spline3.Spline3Interpolate(UBound(TblКоэффициенты, 2) + 1, TblКоэффициенты, ПромКоорХ)
        'ВыходнойМассив(5) = ЗначНаПромКоор
        ''на координате 240
        'ВыходнойМассив(6) = ВсеПараметрыПодсчета("dР310-3-" & Пояс.ToString).dblНакопленноеЗначение

        Return ВыходнойМассив
    End Function

    Public Shared Sub PlotSurface(ByVal СтатическоеДавление As Double)
        'Dim ПолеДавлений As Bicubic
        'Dim ВходнойМассив(,), ВыходнойМассив(,) As Double
        ''шаг по оси Х 4.5 мм от 0 до 27 мм значит 6 промежутков, размерность массива 6 -> 7 точек
        'Dim Высота As Double = 27
        'Dim Ширина As Double = 280
        'Dim I, J As Integer
        'Dim xInit, dtX, xFinal As Double

        ''по У координаты на которых известны значения распределены так
        ''0, 40, 140, 240, 280
        ''растянем через кратный промежуток 40, тогда получим координаты
        ''0, 40, 80, 120, 160, 200, 240, 280
        ''шаг по оси Y 40 мм от 0 до 280 мм значит 7 промежутков, размерность  массива 7 ->  8 точек
        ''где 1 там СтатическоеДавление
        ''где 0 надо вычислить
        ''где 2 в принципе известно - это значения 1 и 3 зонда
        'ВходнойМассив = New Double(,) {{1, 1, 1, 1, 1, 1, 1, 1}, _
        '                                {1, 2, 0, 0, 0, 0, 2, 1}, _
        '                                {1, 2, 0, 0, 0, 0, 2, 1}, _
        '                                {1, 2, 0, 0, 0, 0, 2, 1}, _
        '                                {1, 2, 0, 0, 0, 0, 2, 1}, _
        '                                {1, 2, 0, 0, 0, 0, 2, 1}, _
        '                                {1, 1, 1, 1, 1, 1, 1, 1}}
        ''занести давление у стенок
        'For J = 0 To 7
        '    ВходнойМассив(0, J) = СтатическоеДавление
        '    ВходнойМассив(6, J) = СтатическоеДавление
        'Next
        'For I = 0 To 6
        '    ВходнойМассив(I, 0) = СтатическоеДавление
        '    ВходнойМассив(I, 7) = СтатическоеДавление
        'Next
        'For I = 1 To 5
        '    For J = 1 To 6
        '        'здесь I индекс пояса, J в выходном массиве индекс элемента
        '        ВходнойМассив(I, J) = ДостроитьПолеДавленийНаПромежуточных(I, СтатическоеДавление)(J)
        '    Next
        'Next
        ''растянуть и отдать на график
        'ПолеДавлений = New Bicubic(ВходнойМассив, 0, Высота, 0, Ширина)
        'ВыходнойМассив = ПолеДавлений.OutputArray

        ''вначале матрицу интегрируем сверху вниз (координаты расположены равномерно) для получения 8 проинтегрированных точек
        ''на координатах InputX = New Double() {0, 40, 80, 120, 160, 200, 240, 280}
        'Dim InputY(6) As Double
        'Dim ДавленияНаКоординатах(7) As Double
        'For J = 0 To 7
        '    For I = 0 To 6
        '        InputY(I) = ВходнойМассив(I, J)
        '    Next

        '    dtX = 4.5
        '    xInit = InputY(0) '0
        '    xFinal = InputY(UBound(InputY)) 'ЧислоПромежутков
        '    arrТемперГаза = SignalProcessing.Integrate(InputY, dtX, xInit, xFinal)
        '    ДавленияНаКоординатах(J) = arrТемперГаза(UBound(arrТемперГаза)) / (27 + dtX)
        'Next
        ''затем интегрируем полученные точки (координаты расположены равномерно)
        'Dim dblИнтегральнаяПоПолюДавлений As Double
        'dtX = 40
        'xInit = ДавленияНаКоординатах(0) '0
        'xFinal = ДавленияНаКоординатах(UBound(ДавленияНаКоординатах)) 'ЧислоПромежутков
        'arrТемперГаза = SignalProcessing.Integrate(ДавленияНаКоординатах, dtX, xInit, xFinal)
        'dblИнтегральнаяПоПолюДавлений = arrТемперГаза(UBound(arrТемперГаза)) / (280 + dtX)

        'Dim intТочность As Integer = 3
        ''CWGraph3D2.Plot3DSimpleSurface(ВыходнойМассив)
        ''lblИнтегрПоДавлению = Round(dblИнтегральнаяПоПолюДавлений, intТочность).ToString
        ''Рв_вх_абс_полн_ср = dblИнтегральнаяПоПолюДавлений
        ''lblСтатичесокеДавление = Round(СтатическоеДавление, intТочность).ToString

        ' ''Dim data(0 To 27, 0 To 280) As Double
        ' ''For i As Integer = 0 To 27
        ' ''    For j As Integer = 0 To 280
        ' ''        data(i, j) = Sin(i * Math.PI) '0.314)
        ' ''    Next j
        ' ''Next i
        ' ''CWGraph3D2.Plot3DSimpleSurface(data)
        ' ''CWGraph3D1.Plots(1).Style = cwSurface
        ' ''CWGraph3D1.Plots(1).ColorMapStyle = cwShaded
    End Sub


End Class
