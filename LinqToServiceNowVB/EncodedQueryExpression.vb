Imports System.Linq.Expressions

Public Class EncodedQueryExpression

    Private _query As String = String.Empty

    Public Property ContinuationOperator As Utilities.ContinuationOperator
    Public Property FieldName As String
    Public Property [Operator] As Utilities.RepoExpressionType
    Public Property FieldValue As String
    Public Property IsNegated As Boolean
    Public Property EncodedQuery As String
    Public Property [Expression] As Expression

    Public ReadOnly Property [Value] As String
        Get
            If String.IsNullOrEmpty(_query) Then
                Return EncodedQuery()
            Else
                Return _query
            End If
        End Get
    End Property

    Public ReadOnly Property IsNameOrValueMissing As Boolean
        Get
            Return String.IsNullOrEmpty(Me.FieldName) Or String.IsNullOrEmpty(Me.FieldValue)
        End Get
    End Property

    Public ReadOnly Property HasValue As Boolean
        Get
            Return Not String.IsNullOrEmpty(EncodeValue())
        End Get
    End Property

    Public Overrides Function ToString() As String

        Return EncodeValue()

    End Function

    Private Sub GetTerms()

        If Me.Expression.NodeType = ExpressionType.Call Then
            Dim methodCall As MethodCallExpression = CType(Me.Expression, MethodCallExpression)
            If methodCall.Object IsNot Nothing Then
                Me.FieldName = GetFieldName(methodCall.Object)
                Me.FieldValue = GetFieldValue(methodCall.Arguments(0))
            Else
                SetMethodValues(Me.Expression)
            End If
        Else
            SetValues(Me.Expression)
        End If

    End Sub

    Private Function BuildExpression() As String
        GetTerms()

        If String.IsNullOrEmpty(Me.FieldName) _
            Or String.IsNullOrEmpty(Me.FieldValue) _
            Or String.IsNullOrEmpty(Me.Operator) Then
            If String.IsNullOrEmpty(Me.EncodedQuery) Then
                Return String.Empty
            Else
                Return Me.EncodedQuery
            End If
        End If

        Dim oper As String = Utilities.GetRepoExpressionType(Me.Operator)
        If Me.IsNegated Then
            oper = Utilities.NegateRepoExpressionType(Me.Operator)
        End If

        Return String.Format("{0}{1}{2}", Me.FieldName, oper, Me.FieldValue)
    End Function

    Private Function EncodeValue() As String

        _query = BuildExpression()

        If Not String.IsNullOrEmpty(EncodedQuery) AndAlso Not EncodedQuery.EndsWith("NQ") Then
            _query = Utilities.GetContinuationOperator(ContinuationOperator) & _query
        End If

        If Not String.IsNullOrEmpty(_query) Then
            _query = EncodedQuery & _query
        End If

        Return _query
    End Function

    Private Function GetFieldName(ByVal expr As Expression) As String

        Dim fieldname As String = ""

        If expr.NodeType = ExpressionType.Constant Then
            Throw New Exception("left side of an expression cannot be a constant")
        End If

        If expr.NodeType = ExpressionType.MemberAccess Then
            fieldname = Utilities.GetPropertyName(expr)
        End If

        If expr.NodeType = ExpressionType.Call AndAlso CType(expr, MethodCallExpression).Method.Name = "Parse" Then
            fieldname = Utilities.GetPropertyName(CType(expr, MethodCallExpression).Arguments(0))
        End If

        Return fieldname

    End Function

    Private Function GetFieldValue(ByVal expr As Expression) As String

        Dim fieldvalue As String = ""

        If expr.NodeType = ExpressionType.Constant Then
            fieldvalue = expr.ToString().Replace("""", "")
        End If

        If expr.NodeType = ExpressionType.MemberAccess Then
            fieldvalue = Utilities.GetPropertyName(expr)
        End If

        If expr.NodeType = ExpressionType.Call Then
            Dim methodCall As MethodCallExpression = CType(expr, MethodCallExpression)
            If methodCall.Method.Name = "Parse" Then
                fieldvalue = methodCall.Arguments(0).ToString().Replace("""", "")
            Else
                fieldvalue = Expression.Lambda(expr).Compile().DynamicInvoke()
            End If
        End If

        If expr.NodeType = ExpressionType.Convert Then
            fieldvalue = GetFieldValue(CType(expr, UnaryExpression).Operand)
        End If

        If expr.NodeType = ExpressionType.NewArrayInit Then
            fieldvalue = String.Join(",", CType(expr, NewArrayExpression).Expressions.Select(Function(o) o.ToString().Replace("""", "")).ToArray())
        End If

        Return fieldvalue

    End Function

    Private Sub SetMethodValues(ByVal methodCall As MethodCallExpression)

        If methodCall.Arguments.Count > 1 Then
            If methodCall.Object IsNot Nothing Then
                Me.FieldName = GetFieldName(methodCall.Object)
                Me.FieldValue = GetFieldValue(methodCall.Arguments(0))
            ElseIf methodCall.Arguments(1).NodeType = ExpressionType.MemberAccess Then
                Me.FieldName = GetFieldName(methodCall.Arguments(1))
                Me.FieldValue = GetFieldValue(methodCall.Arguments(0))
            Else
                Me.FieldName = GetFieldName(methodCall.Arguments(0))
                Me.FieldValue = GetFieldValue(methodCall.Arguments(1))
            End If
        Else
            If methodCall.Arguments(0).NodeType = ExpressionType.Call Then
                SetMethodValues(methodCall.Arguments(0))
            End If
            If methodCall.Arguments(0).NodeType = ExpressionType.Constant Then 'flip expression
                Me.FieldValue = GetFieldValue(methodCall)
            End If
            If methodCall.Arguments(0).NodeType = ExpressionType.MemberAccess Then
                Me.FieldName = GetFieldName(methodCall)
            End If
        End If

    End Sub

    Private Sub SetBinaryValues(ByVal binExpr As BinaryExpression)

        FlipOperator(binExpr.Left)
        SetValues(binExpr.Left)
        If IsNameOrValueMissing Then
            SetValues(binExpr.Right)
        End If

    End Sub

    Private Sub SetValues(ByVal expr As Expression)

        If expr.NodeType = ExpressionType.Constant Then
            Me.FieldValue = GetFieldValue(expr)
        ElseIf expr.NodeType = ExpressionType.MemberAccess Then 'flip expression
            Me.FieldName = GetFieldName(expr)
        ElseIf expr.NodeType = ExpressionType.Call Then
            SetMethodValues(CType(expr, MethodCallExpression))
        ElseIf expr.NodeType = ExpressionType.Convert Then
            SetValues(CType(expr, UnaryExpression).Operand)
        Else
            SetBinaryValues(CType(expr, BinaryExpression))
        End If

    End Sub

    Private Sub FlipOperator(ByVal expr As Expression)

        If expr.NodeType = ExpressionType.Constant Then 'flip expression
            Me.Operator = Utilities.FlipRepoExpressionType(Me.Operator)
        ElseIf expr.NodeType = ExpressionType.Call Then
            Dim methodCall As MethodCallExpression = CType(expr, MethodCallExpression)
            If methodCall.Arguments(0).NodeType = ExpressionType.Constant Then 'flip expression
                Me.Operator = Utilities.FlipRepoExpressionType(Me.Operator)
            End If
            If methodCall.Arguments(0).NodeType = ExpressionType.Call Then 'flip expression
                FlipOperator(methodCall.Arguments(0))
            End If
        End If

    End Sub

End Class
