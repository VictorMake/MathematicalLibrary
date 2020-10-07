Imports System.Windows.Forms

Public Class LinearAlgebra

    Dim computeClicked As Integer = 0

    Public Sub New()
        'Здесь может быть инициализация массивов
    End Sub

    'Read 2D Matrix from the mainForm.cs panel.
    Private Function Read2DMatrix(ByVal matrixData As String) As Double(,)
        Dim rowIndex, columnIndex As Integer
        Dim splitRows As String() = Text.RegularExpressions.Regex.Split(matrixData, ";", Text.RegularExpressions.RegexOptions.Multiline)
        Dim splitDataByColumns As String()
        Dim matrix As Double(,) = New Double(splitRows.Length - 2, splitRows.Length - 2) {}

        Dim numFirstRowElements As Integer = 0
        'store rows of Matrix in to splitRows string array.

        'Memory Allocation for matrix.

        For rowIndex = 0 To (splitRows.Length - 2)
            splitDataByColumns = Text.RegularExpressions.Regex.Split(splitRows(rowIndex), ",")
            If (rowIndex = 0) Then
                numFirstRowElements = splitDataByColumns.Length
                matrix = New Double(splitRows.Length - 2, splitDataByColumns.Length - 1) {}
            Else
                If numFirstRowElements <> splitDataByColumns.Length Then
                    Throw New ArgumentException(String.Concat("Number of columns is incorrect for Row ", (rowIndex + 1)))
                End If
            End If
            For columnIndex = 0 To (splitDataByColumns.Length - 1)
                matrix(rowIndex, columnIndex) = Convert.ToDouble(splitDataByColumns(columnIndex))
            Next columnIndex
        Next rowIndex
        Return matrix
    End Function

    'Read 1D Matrix.
    Private Function Read1DMatrix(ByVal matrixData As String) As Double()
        Dim splitRows As String() = Text.RegularExpressions.Regex.Split(matrixData, ";", Text.RegularExpressions.RegexOptions.Multiline)
        Dim matrix As Double() = New Double(splitRows.Length - 2) {}
        Dim i As Integer
        For i = 0 To (splitRows.Length - 2)
            matrix(i) = Convert.ToDouble(splitRows(i))
        Next i
        Return matrix
    End Function

    Private varmatrixADataTextBox As String
    Public Property matrixADataTextBox() As String
        Get
            Return varmatrixADataTextBox
        End Get
        Set(ByVal value As String)
            varmatrixADataTextBox = value
        End Set
    End Property

    Private varmatrixBDataTextBox As String
    Public Property matrixBDataTextBox() As String
        Get
            Return varmatrixBDataTextBox
        End Get
        Set(ByVal value As String)
            varmatrixBDataTextBox = value
        End Set
    End Property

    Private varmatrixXDataTextBox As String
    Public Property matrixXDataTextBox() As String
        Get
            Return varmatrixXDataTextBox
        End Get
        Set(ByVal value As String)
            varmatrixXDataTextBox = value
        End Set
    End Property


    Public Enum EnumNormTypeComboBox
        norm2
        norm1
        Fnorm
        infnorm
    End Enum

    Public normTypeComboBox As EnumNormTypeComboBox = EnumNormTypeComboBox.norm2

    Public Enum EnumOperationsComboBox
        Determinant
        Trace
        Rank
        ConditionNumber
        Inverse
        PseudoInverse
        Norm
        TestPositiveDefinite
        Transpose
        CholeskyFactorization
        Eigen_Values_and_Eigen_Vectors
        LUDecomposition
        QRFactorization
        SolveLinearEquations_AxB
        SingularValueDecomposition
    End Enum

    'Dim name As String = _
    '    System.Enum.GetName(GetType(EnumOperationsComboBox), 0)

    Public operationsComboBox As EnumOperationsComboBox = EnumOperationsComboBox.Determinant

    Private varOperations As String
    Public Property Operations() As String
        Get
            Return varOperations
        End Get
        Set(ByVal value As String)
            varOperations = value
        End Set
    End Property

    'create a data set.
    Public data As System.Data.DataSet ' = New DataSet(" ")
    Public table As System.Data.DataTable ' = data.Tables.Add(" ") 'Attach a table to data set.


    'Public Sub Compute()
    '    data = New DataSet(" ")
    '    table = data.Tables.Add(" ") 'Attach a table to data set.
    '    Dim value1() As String = New String(0) {}
    '    Dim determinant As Double = 0

    '    value1(0) = String.Format("{0:F2}", determinant)
    '    table.TableName = "Determinant"
    '    table.Columns.Add(" ")
    '    table.Rows.Add(value1)
    '    varOperations = "   Determinant of Matrix A"
    '    'operationsDataGrid.DataSource = data

    'End Sub

    Public Sub Compute()
        Dim i, j As Integer
        Dim dataString As String()
        Dim value1 As String() = New String(0) {}
        Dim matrixA As Double(,) = {}
        Dim normTypeValue As NationalInstruments.Analysis.Math.NormType

        computeClicked = 1 'set it to one when comput button is clicked.
        value1.Initialize()

        'create a data set.
        'Dim data As System.Data.DataSet = New DataSet(" ")
        'Dim table As System.Data.DataTable = data.Tables.Add(" ") 'Attach a table to data set.

        'create a data set.
        data = New DataSet(" ")
        table = data.Tables.Add(" ") 'Attach a table to data set.

        Try
            matrixA = Read2DMatrix(varmatrixADataTextBox) 'read Matrix A.

            If matrixA.GetLength(0) = 0 OrElse matrixA.GetLength(1) = 0 Then
                Throw New ArgumentException("Matrix A has incorrect number of rows or columns")
            End If

            Select Case operationsComboBox
                Case EnumOperationsComboBox.Determinant '0 Determinant
                    Dim determinant As Double = NationalInstruments.Analysis.Math.LinearAlgebra.Determinant(matrixA)
                    'determinant = Analysis.Math.LinearAlgebra.Determinant(matrixA)
                    value1(0) = String.Format("{0:F2}", determinant)
                    table.TableName = "Determinant"
                    table.Columns.Add(" ")
                    table.Rows.Add(value1)
                    varOperations = "   Determinant of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.Trace '1 'Trace
                    Dim trace As Double = 0

                    trace = NationalInstruments.Analysis.Math.LinearAlgebra.Trace(matrixA)
                    table.TableName = "Trace"
                    value1(0) = String.Format("{0:F2}", trace)
                    table.Columns.Add("  ")
                    table.Rows.Add(value1)
                    varOperations = "   Trace of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.Rank '2 'Rank
                    Dim rank As Integer

                    rank = NationalInstruments.Analysis.Math.LinearAlgebra.Rank(matrixA, -1)
                    table.TableName = "Rank"
                    value1(0) = String.Format("{0:F2}", rank)
                    table.Columns.Add("  ")
                    table.Rows.Add(value1)
                    varOperations = "   Rank of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.ConditionNumber '3 'Condition Number.
                    Dim conditionNumber As Double

                    Select Case normTypeComboBox
                        Case EnumNormTypeComboBox.norm2 '0 '2-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.TwoNorm
                        Case EnumNormTypeComboBox.norm1 '1 '1-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.OneNorm
                        Case EnumNormTypeComboBox.Fnorm '2 'F-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.FrobeniusNorm
                        Case EnumNormTypeComboBox.infnorm '3 'I-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.InfiniteNorm
                    End Select

                    conditionNumber = NationalInstruments.Analysis.Math.LinearAlgebra.ConditionNumber(matrixA, normTypeValue)
                    table.TableName = "ConditionNumber"
                    value1(0) = String.Format("{0:F2}", conditionNumber)
                    table.Columns.Add("  ")
                    table.Rows.Add(value1)
                    varOperations = "   Condition Number of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.Inverse '4 'Inverse Matrix
                    Dim inverseMatrix As Double(,)

                    If (matrixA.GetLength(0) = matrixA.GetLength(1)) Then
                        inverseMatrix = NationalInstruments.Analysis.Math.LinearAlgebra.Inverse(matrixA)
                        dataString = New String(matrixA.GetLength(0) - 1) {}
                        table.TableName = "InverseMatrix"
                        For i = 0 To (matrixA.GetLength(0) - 1)
                            table.Columns.Add()
                        Next i
                        For i = 0 To (matrixA.GetLength(0) - 1)
                            For j = 0 To (matrixA.GetLength(0) - 1)
                                dataString(j) = String.Format("{0:F2}", inverseMatrix(i, j))
                            Next j
                            table.Rows.Add(dataString)
                        Next i
                        varOperations = "   Inverse of Matrix A"
                        'operationsDataGrid.DataSource = data
                    Else
                        MessageBox.Show("Please enter a square matrix to perform this operation")
                    End If
                Case EnumOperationsComboBox.PseudoInverse '5 'Pseudo Inverse Matrix.
                    Dim pseudoInverse As Double(,) = NationalInstruments.Analysis.Math.LinearAlgebra.PseudoInverse(matrixA, -1)
                    dataString = New String(pseudoInverse.GetLength(1) - 1) {}
                    table.TableName = " Pseudo Inverse"
                    For i = 0 To (pseudoInverse.GetLength(1) - 1)
                        table.Columns.Add()
                    Next i
                    For i = 0 To (pseudoInverse.GetLength(0) - 1)
                        For j = 0 To (pseudoInverse.GetLength(1) - 1)
                            dataString(j) = String.Format("{0:F2}", pseudoInverse(i, j))
                        Next j
                        table.Rows.Add(dataString)
                    Next i
                    varOperations = "   Pseudo Inverse of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.Norm '6 'Norm of the Marix.
                    Dim norm As Double

                    Select Case normTypeComboBox
                        Case EnumNormTypeComboBox.norm2 '0 '2-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.TwoNorm
                        Case EnumNormTypeComboBox.norm1 '1 '1-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.OneNorm
                        Case EnumNormTypeComboBox.Fnorm '2 'F-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.FrobeniusNorm
                        Case EnumNormTypeComboBox.infnorm '3 'I-norm.
                            normTypeValue = NationalInstruments.Analysis.Math.NormType.InfiniteNorm
                    End Select
                    norm = NationalInstruments.Analysis.Math.LinearAlgebra.Norm(matrixA, normTypeValue)
                    table.TableName = "Norm"
                    value1(0) = String.Format("{0:F2}", norm)
                    table.Columns.Add("  ")
                    table.Rows.Add(value1)
                    varOperations = "   Norm of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.TestPositiveDefinite '7 'Is PositiveDefinite?
                    Dim isPositiveDefinite As Boolean = False

                    If (matrixA.GetLength(0) = matrixA.GetLength(1)) Then
                        isPositiveDefinite = NationalInstruments.Analysis.Math.LinearAlgebra.IsPositiveDefinite(matrixA)
                        table.TableName = "isPositiveDefinite?"
                        If (isPositiveDefinite = False) Then
                            value1(0) = "False"
                        Else
                            value1(0) = "True"
                        End If
                        table.Columns.Add("  ")
                        table.Rows.Add(value1)
                        varOperations = "Test Positive Definite of Matrix A"
                        'operationsDataGrid.DataSource = data
                    Else
                        MessageBox.Show("Please enter a square matrix to perform this operation")
                    End If
                Case EnumOperationsComboBox.Transpose ' 8 'Transpose
                    Dim transpose As Double(,) = NationalInstruments.Analysis.Math.LinearAlgebra.Transpose(matrixA)
                    dataString = New String(transpose.GetLength(1) - 1) {}
                    table.TableName = " Transpose"
                    For i = 0 To (transpose.GetLength(1) - 1)
                        table.Columns.Add()
                    Next i
                    For i = 0 To (transpose.GetLength(0) - 1)
                        For j = 0 To (transpose.GetLength(1) - 1)
                            dataString(j) = String.Format("{0:F2}", transpose(i, j))
                        Next j
                        table.Rows.Add(dataString)
                    Next i
                    varOperations = "   Transpose of Matrix A"
                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.CholeskyFactorization '9 'Cholesky Factorization.
                    Dim choleskyFactorization As Double(,)

                    If ((matrixA.GetLength(0)) = (matrixA.GetLength(1))) Then

                        If (NationalInstruments.Analysis.Math.LinearAlgebra.IsPositiveDefinite(matrixA) = False) Then
                            MessageBox.Show("Input matrix is non singular.", "Error")
                        Else
                            choleskyFactorization = New Double(matrixA.GetLength(1) - 1, matrixA.GetLength(0) - 1) {}
                            choleskyFactorization = NationalInstruments.Analysis.Math.LinearAlgebra.CholeskyFactorization(matrixA)
                            dataString = New String(choleskyFactorization.GetLength(1) - 1) {}
                            table.TableName = " CholeskyFactorization"
                            For i = 0 To (choleskyFactorization.GetLength(1) - 1)
                                table.Columns.Add()
                            Next i
                            For i = 0 To (choleskyFactorization.GetLength(0) - 1)
                                For j = 0 To (choleskyFactorization.GetLength(1) - 1)
                                    dataString(j) = String.Format("{0:F2}", choleskyFactorization(i, j))
                                Next j
                                table.Rows.Add(dataString)
                            Next i

                            varOperations = "Cholesky Factorization of Matrix A"
                            'operationsDataGrid.DataSource = data
                        End If
                    Else
                        MessageBox.Show("Please enter a square and non singular matrix to perform this operation")
                    End If
                Case EnumOperationsComboBox.Eigen_Values_and_Eigen_Vectors '10 'Eigen Values and Eigen Vectors.
                    Dim eigenValues As ComplexDouble() = {}
                    Dim eigenVectors As ComplexDouble(,) = {}

                    If (matrixA.GetLength(0) = matrixA.GetLength(1)) Then
                        eigenValues = New ComplexDouble(matrixA.GetLength(0) - 1) {}
                        eigenValues = NationalInstruments.Analysis.Math.LinearAlgebra.GeneralEigenValueVector(matrixA, eigenVectors)
                        'Display Eigen Values in a data set table.
                        dataString = New String(0) {}
                        table.TableName = " Eigen Values"
                        table.Columns.Add(" ")

                        For i = 0 To (eigenValues.GetLength(0) - 1)
                            dataString(0) = String.Format("{0:F2}", eigenValues(i))
                            table.Rows.Add(dataString)
                        Next i

                        varOperations = "Eigen Values and Eigen Vectors of Matrix A"
                        'operationsDataGrid.DataSource = data
                        'Display eigen vectors in another table.
                        Dim table2 As System.Data.DataTable = data.Tables.Add("Eigen Vectors ")
                        dataString = New String(eigenVectors.GetLength(1) - 1) {}

                        For i = 0 To (eigenVectors.GetLength(1) - 1)
                            table2.Columns.Add()
                        Next i
                        For i = 0 To (eigenVectors.GetLength(0) - 1)
                            For j = 0 To (eigenVectors.GetLength(1) - 1)
                                dataString(j) = String.Format("{0:F2}", eigenVectors(i, j))
                            Next j
                            table2.Rows.Add(dataString)
                        Next i
                        'operationsDataGrid.DataSource = data
                    Else
                        MessageBox.Show("Please enter a square matrix to perform this operation")
                    End If
                Case EnumOperationsComboBox.LUDecomposition '11 'LU Factorization.
                    Dim sign As Integer
                    Dim permutationVector As Integer() = {}
                    Dim luFactorize As Double(,)

                    If (matrixA.GetLength(0) = matrixA.GetLength(1)) Then
                        luFactorize = NationalInstruments.Analysis.Math.LinearAlgebra.LUFactorization(matrixA, permutationVector, sign)
                        dataString = New String(luFactorize.GetLength(1) - 1) {}
                        table.TableName = "LUFactorization"
                        For i = 0 To (luFactorize.GetLength(1) - 1)
                            table.Columns.Add()
                        Next i
                        For i = 0 To (luFactorize.GetLength(0) - 1)
                            For j = 0 To (luFactorize.GetLength(1) - 1)
                                dataString(j) = String.Format("{0:F2}", luFactorize(i, j))
                            Next j
                            table.Rows.Add(dataString)
                        Next i
                        varOperations = "LU Factorization of Matrix A"
                        'operationsDataGrid.DataSource = data
                    Else
                        MessageBox.Show("Please enter a square matrix to perform this operation")
                    End If
                Case EnumOperationsComboBox.QRFactorization '12 'QR Factorization.
                    Dim qMatrix As Double(,) = {}
                    Dim rMatrix As Double(,) = {}

                    NationalInstruments.Analysis.Math.LinearAlgebra.QRFactorization(matrixA, SizeOption.Economy, rMatrix, qMatrix)
                    'Display Q Matrix.
                    dataString = New String(qMatrix.GetLength(1) - 1) {}
                    table.TableName = "Q Matrix"
                    For i = 0 To (qMatrix.GetLength(1) - 1)
                        table.Columns.Add()
                    Next i
                    For i = 0 To (qMatrix.GetLength(0) - 1)
                        For j = 0 To (qMatrix.GetLength(1) - 1)
                            dataString(j) = String.Format("{0:F2}", qMatrix(i, j))
                        Next j
                        table.Rows.Add(dataString)
                    Next i
                    varOperations = "QR Factorization of Matrix A"
                    'operationsDataGrid.DataSource = data
                    'Display R Matrix.
                    Dim table3 As System.Data.DataTable = data.Tables.Add("R Matrix ")
                    dataString = New String(rMatrix.GetLength(1) - 1) {}
                    For i = 0 To (rMatrix.GetLength(1) - 1)
                        table3.Columns.Add()
                    Next i
                    For i = 0 To (rMatrix.GetLength(0) - 1)
                        For j = 0 To (rMatrix.GetLength(1) - 1)
                            dataString(j) = String.Format("{0:F2}", rMatrix(i, j))
                        Next j
                        table3.Rows.Add(dataString)
                    Next i

                    'operationsDataGrid.DataSource = data
                Case EnumOperationsComboBox.SolveLinearEquations_AxB '13 'Solve Linear Equations.
                    Dim arrayB As Double()
                    Dim solutionMatrix As Double()
                    Dim rowStringArray As String()

                    If (matrixA.GetLength(0) = matrixA.GetLength(1)) Then
                        arrayB = Read1DMatrix(varmatrixBDataTextBox)
                        solutionMatrix = NationalInstruments.Analysis.Math.LinearAlgebra.SolveLinearEquationsSingleRightHand(matrixA, MatrixType.General, arrayB)
                        dataString = New String(0) {}
                        rowStringArray = New String(solutionMatrix.Length - 1) {}
                        table.TableName = "Solution x to Ax = B"
                        table.Columns.Add(" ")

                        For i = 0 To (solutionMatrix.GetLength(0) - 1)
                            dataString(0) = String.Format("{0:F2}", solutionMatrix(i))
                            rowStringArray(i) = String.Format("{0:F2}", solutionMatrix(i))
                            table.Rows.Add(dataString)
                        Next i
                        Dim sep As String = vbNewLine
                        varmatrixXDataTextBox = String.Join(sep, rowStringArray)
                        varOperations = " Solution x to Ax = B"
                        'operationsDataGrid.DataSource = data
                    Else
                        MessageBox.Show("Please enter a square matrix to perform this operation")
                    End If
                Case EnumOperationsComboBox.SingularValueDecomposition '14 'Singular Value Decomposition. 
                    Dim uMatrix As Double(,) = {}
                    Dim vMatrix As Double(,) = {}
                    Dim singularValues As Double() =
                    NationalInstruments.Analysis.Math.LinearAlgebra.SvdFactorization(matrixA, SizeOption.Economy, uMatrix, vMatrix)
                    'Display of singular values.
                    dataString = New String(0) {}

                    table.TableName = "Singular Values"
                    table.Columns.Add(" ")

                    For i = 0 To (singularValues.GetLength(0) - 1)
                        dataString(0) = String.Format("{0:F2}", singularValues(i))
                        table.Rows.Add(dataString)
                    Next i
                    varOperations = "Singular Value Decomposition of MatrixA"
                    'operationsDataGrid.DataSource = data
                    'To display U matrix.
                    Dim table4 As DataTable = data.Tables.Add("SVD of A - U ")
                    dataString = New String(uMatrix.GetLength(1) - 1) {}

                    For i = 0 To (uMatrix.GetLength(1) - 1)
                        table4.Columns.Add()
                    Next i
                    For i = 0 To (uMatrix.GetLength(0) - 1)
                        For j = 0 To (uMatrix.GetLength(1) - 1)
                            dataString(j) = String.Format("{0:F2}", uMatrix(i, j))
                        Next j
                        table4.Rows.Add(dataString)
                    Next i
                    'Display V - Matrix.
                    Dim table5 As DataTable = data.Tables.Add("SVD of A - V ")
                    dataString = New String(vMatrix.GetLength(1) - 1) {}

                    For i = 0 To (vMatrix.GetLength(1) - 1)
                        table5.Columns.Add()
                    Next i
                    For i = 0 To (vMatrix.GetLength(0) - 1)
                        For j = 0 To (vMatrix.GetLength(1) - 1)
                            dataString(j) = String.Format("{0:F2}", vMatrix(i, j))
                        Next j
                        table5.Rows.Add(dataString)
                    Next i

                    'operationsDataGrid.DataSource = data
            End Select
        Catch exp As Exception
            MessageBox.Show(exp.Message)
        End Try
    End Sub

    'Private Sub operations_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles operationsComboBoxChanged
    '    If operationsComboBox = 3 Or operationsComboBox = 6 Then
    '        normTypeComboBox.Enabled = True
    '        normTypeLabel.Enabled = True
    '    Else
    '        normTypeComboBox.Enabled = False
    '        normTypeLabel.Enabled = False
    '    End If
    '    If (operationsComboBox = 13) Then
    '        matrixBDataTextBox.Enabled = True
    '        matrixBLabel.Enabled = True
    '        matrixXLabel.Enabled = True
    '        matrixXDataTextBox.Enabled = True
    '        linearEquationLabel.Enabled = True
    '    Else
    '        matrixBDataTextBox.Enabled = False
    '        matrixBLabel.Enabled = False
    '        matrixXLabel.Enabled = False
    '        matrixXDataTextBox.Enabled = False
    '        linearEquationLabel.Enabled = False
    '    End If
    '    If (computeClicked = 1) Then
    '        computeButton.PerformClick()
    '    End If
    'End Sub


End Class
