''' <summary>
''' Bсе вызовы ReDimPreserve X(N) заменили на Re.DimPreserve(X,N), а ReDim_X(N) заменили на Re.Dim(X,N).
''' </summary>
Public Class Re
    ''' <summary>
    ''' Очищение массива
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Source"></param>
    ''' <param name="I"></param>
    ''' <param name="J"></param>
    Public Shared Sub [Dim](Of T)(ByRef Source As T(,), ByVal I As Integer, ByVal J As Integer)
        Source = New T(I, J) {}
    End Sub
    ''' <summary>
    ''' Очищение массива
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Source"></param>
    ''' <param name="I"></param>
    ''' <param name="J"></param>
    ''' <param name="K"></param>
    Public Shared Sub [Dim](Of T)(ByRef Source As T(,,), ByVal I As Integer, ByVal J As Integer, ByVal K As Integer)
        Source = New T(I, J, K) {}
    End Sub
    ''' <summary>
    ''' Очищение массива
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Source"></param>
    ''' <param name="I"></param>
    ''' <param name="J"></param>
    ''' <param name="K"></param>
    ''' <param name="L"></param>
    Public Shared Sub [Dim](Of T)(ByRef Source As T(,,,), ByVal I As Integer, ByVal J As Integer, ByVal K As Integer, ByVal L As Integer)
        Source = New T(I, J, K, L) {}
    End Sub

    ''' <summary>
    ''' Очищение массива
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Source"></param>
    ''' <param name="HighBound"></param>
    Public Shared Sub [Dim](Of T)(ByRef Source As T(), HighBound As Integer)
        DimPreserve(Source, HighBound)
        Array.Clear(Source, 0, HighBound + 1)
    End Sub

    ''' <summary>
    ''' Изменение размера массива с сохранением содержимого
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Source"></param>
    ''' <param name="HighBound"></param>
    Public Shared Sub DimPreserve(Of T)(ByRef Source As T(), HighBound As Integer)
        Array.Resize(Source, HighBound + 1)
    End Sub

    ''' <summary>
    ''' Изменение размера массива с сохранением содержимого
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Source"></param>
    ''' <param name="I"></param>
    ''' <param name="J"></param>
    Public Shared Sub DimPreserve(Of T)(ByRef Source As T(,), ByVal I As Integer, ByVal J As Integer)
        Dim tempVar As T(,) = New T(I, J) {}

        If Source IsNot Nothing Then
            For Dimension0 As Integer = 0 To Source.GetLength(0) - 1
                Dim CopyLength As Integer = Math.Min(Source.GetLength(1), tempVar.GetLength(1))

                For Dimension1 As Integer = 0 To CopyLength - 1
                    tempVar(Dimension0, Dimension1) = Source(Dimension0, Dimension1)
                Next
            Next
        End If

        Source = tempVar
    End Sub

    ' Первоисточник
    'Friend Sub ReDim(ByRef Source As Double(,), I As Integer, J As Integer)
    '    ReDim_Source(I, J)
    'End Sub
    ' Сконвертировано в C#
    'internal void ReDim(ref Double[,] Source, int I, int J)
    '{
    '	Source = New double[I + 1, J + 1];
    '}
    'Переведено назад в VB
    'Friend Sub [ReDim](ByRef Source As Double(,), ByVal I As Integer, ByVal J As Integer)
    '    Source = New Double(I + 1 - 1, J + 1 - 1) {}
    'End Sub


    ' Первоисточник
    'Friend Sub DimPreserve(I As Integer, J As Integer)
    '    Dim y(I, J) As Double

    '    ReDim_y(I, J)
    '    ReDim_Preserve y(I, J)
    'End Sub


    ' Сконвертировано в C#
    'internal void DimPreserve(int I, int J)
    '{
    '	Double[,] y = New Double[I + 1, J + 1];

    '		y = New double[I + 1, J + 1];
    '//INSTANT C# NOTE: The following block reproduces what 'ReDim_Preserve' does behind the scenes in VB:
    '//ORIGINAL LINE: ReDim_Preserve y(I, J)
    '    Double[,] tempVar = New Double[I + 1, J + 1];
    '	If (y!= null)
    '	{
    '		For (int Dimension0 = 0; Dimension0 < y.GetLength(0); Dimension0++)
    '		{
    '			int CopyLength = Math.Min(y.GetLength(1), tempVar.GetLength(1));
    '			For (int Dimension1 = 0; Dimension1 < CopyLength; Dimension1++)
    '			{
    '				tempVar[Dimension0, Dimension1] = y[Dimension0, Dimension1];
    '			}
    '		}
    '	}
    '	y = tempVar;
    '}

    'Переведено назад в VB
    'Friend Sub DimPreserve(ByVal I As Integer, ByVal J As Integer)
    '    Dim y As Double(,) = New Double(I, J) {}
    '    y = New Double(I, J) {}
    '    Dim tempVar As Double(,) = New Double(I, J) {}

    '    If y IsNot Nothing Then

    '        For Dimension0 As Integer = 0 To y.GetLength(0) - 1
    '            Dim CopyLength As Integer = Math.Min(y.GetLength(1), tempVar.GetLength(1))

    '            For Dimension1 As Integer = 0 To CopyLength - 1
    '                tempVar(Dimension0, Dimension1) = y(Dimension0, Dimension1)
    '            Next
    '        Next
    '    End If

    '    y = tempVar
    'End Sub
End Class
