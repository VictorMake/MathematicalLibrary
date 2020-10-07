Imports System.Math

Public Class Bicubic
    Private ВходнойМассив(,), ВыходнойМассив(,) As Double
    Private ReadOnly OldWidth As Integer  '- старый размер сетки
    Private ReadOnly OldHeight As Integer '- старый размер сетки
    Private ReadOnly NewWidth As Integer  '- новый размер сетки
    Private ReadOnly NewHeight As Integer '- новый размер сетки
    Private ReadOnly mМножитель As Integer
    Private ReadOnly Xstart As Double
    Private ReadOnly Xfinish As Double
    Private ReadOnly Ystart As Double
    Private ReadOnly Yfinish As Double
    Private mВнеДиапазона As Boolean
    Private RangeX, RangeY As Double

    Public Sub New(ByVal InputArray As Double(,), ByVal XIstart As Double, ByVal XIfinish As Double, ByVal YJstart As Double, ByVal YJfinish As Double)
        MyBase.New()
        mМножитель = 10
        ВходнойМассив = InputArray
        OldHeight = UBound(InputArray, 1) + 1
        OldWidth = UBound(InputArray, 2) + 1
        NewHeight = OldHeight * mМножитель
        NewWidth = OldWidth * mМножитель
        Xstart = XIstart
        Xfinish = XIfinish
        Ystart = YJstart
        Yfinish = YJfinish

        InitializeComponent()
    End Sub

    Public Sub New(ByVal InputArray As Double(,), ByVal XIstart As Double, ByVal XIfinish As Double, ByVal YJstart As Double, ByVal YJfinish As Double, ByVal NewHeightOutputArray As Integer, ByVal NewWidthOutputArray As Integer)
        MyBase.New()
        'mМножитель = 10
        ВходнойМассив = InputArray
        OldHeight = UBound(InputArray, 1) + 1
        OldWidth = UBound(InputArray, 2) + 1
        NewHeight = NewHeightOutputArray
        NewWidth = NewWidthOutputArray
        Xstart = XIstart
        Xfinish = XIfinish
        Ystart = YJstart
        Yfinish = YJfinish

        InitializeComponent()
    End Sub

    Public Sub New(ByVal InputArray As Double(,), ByVal XIstart As Double, ByVal XIfinish As Double, ByVal YJstart As Double, ByVal YJfinish As Double, ByVal Множитель As Integer)
        MyBase.New()
        mМножитель = Множитель
        ВходнойМассив = InputArray
        OldHeight = UBound(InputArray, 1) + 1
        OldWidth = UBound(InputArray, 2) + 1
        NewHeight = OldHeight * mМножитель
        NewWidth = OldWidth * mМножитель
        Xstart = XIstart
        Xfinish = XIfinish
        Ystart = YJstart
        Yfinish = YJfinish

        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        'ReDim_ВыходнойМассив(NewHeight - 1, NewWidth - 1)
        Re.Dim(ВыходнойМассив, NewHeight - 1, NewWidth - 1)
        Spline3.BicubicResample(OldWidth, OldHeight, NewWidth, NewHeight, ВходнойМассив, ВыходнойМассив)
        RangeX = Abs(Xfinish - Xstart)
        RangeY = Abs(Yfinish - Ystart)
    End Sub

    Public Function Calculate(ByVal ТекущийХ As Double, ByVal ТекущийY As Double) As Double
        Dim Inew, Jnew As Integer
        Dim ВнеДиапазонаX, ВнеДиапазонаY As Boolean

        If ТекущийХ <= Xstart Then
            Inew = 0
            ВнеДиапазонаX = True
        ElseIf ТекущийХ >= Xfinish Then
            Inew = NewHeight - 1
            ВнеДиапазонаX = True
        Else
            Inew = CInt((NewHeight - 1) / ((RangeX / Abs(ТекущийХ - Xstart))))
        End If

        If ТекущийY <= Ystart Then
            Jnew = 0
            ВнеДиапазонаY = True
        ElseIf ТекущийY >= Yfinish Then
            Jnew = NewWidth - 1
            ВнеДиапазонаY = True
        Else
            Jnew = CInt((NewWidth - 1) / ((RangeY / Abs(ТекущийY - Ystart))))
        End If

        mВнеДиапазона = ВнеДиапазонаX OrElse ВнеДиапазонаY
        Return ВыходнойМассив(Inew, Jnew)
    End Function

    Public ReadOnly Property ВнеДиапазона() As Boolean
        Get
            Return mВнеДиапазона
        End Get
    End Property

    'Private PropertyValue As String

    '' Define the property.
    'Public Property Prop1() As String
    '    Get
    '        ' The Get property procedure is called when the value
    '        ' of a property is retrieved. 
    '        Return PropertyValue
    '    End Get
    '    Set(ByVal Value As String)
    '        ' The Set property procedure is called when the value 
    '        ' of a property is modified. 
    '        ' The value to be assigned is passed in the argument to Set. 
    '        PropertyValue = Value
    '    End Set
    'End Property

    'Public WriteOnly Property InputArray() As Double(,)
    '    Set(ByVal Value As Double(,))
    '        mInputArray = Value
    '    End Set
    'End Property

    'Public WriteOnly Property InputArray() As Double(,)
    '    Set(ByVal Value As Double(,))
    '        mInputArray = Value
    '    End Set
    'End Property
    'Private ВыходнойМассивValue As Double(,)
    Public ReadOnly Property OutputArray() As Double(,)
        Get
            Return ВыходнойМассив
        End Get
    End Property
End Class
