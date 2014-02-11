Imports System.Linq.Expressions

Public Class ExpressionVisitor
    Private _encodedQuery As String = String.Empty

    Public Function VisitExpression(ByVal continuation As Utilities.ContinuationOperator, ByVal expr As Expression) As String
        Return VisitExpression(continuation, expr, False)
    End Function

    Public Function VisitExpression(ByVal continuation As Utilities.ContinuationOperator, ByVal expr As Expression, ByVal neg As Boolean) As String

        If expr.NodeType = ExpressionType.Call Then
            Dim methodCall As MethodCallExpression = CType(expr, MethodCallExpression)
            If methodCall.Method.Name = "Contains" And methodCall.Arguments.Count > 1 Then
                VisitContainsExpression(continuation, methodCall, neg)
            Else
                VisitSimpleMethodCall(continuation, methodCall, neg)
            End If
        ElseIf expr.NodeType = ExpressionType.Not Then
            Dim unaryExpr As UnaryExpression = CType(expr, UnaryExpression)
            VisitExpression(continuation, unaryExpr.Operand, True)
        Else
            VisitBinaryExpression(continuation, CType(expr, BinaryExpression), neg)
        End If

        Return _encodedQuery

    End Function

    Private Sub VisitBinaryExpression(ByVal continuation As Utilities.ContinuationOperator, ByVal binExpr As BinaryExpression, ByVal neg As Boolean)

        Dim AndOperators As ExpressionType() = {ExpressionType.And, ExpressionType.AndAlso}

        Dim OrOperators As ExpressionType() = {ExpressionType.Or, ExpressionType.OrElse}

        Dim binOperators As ExpressionType() = AndOperators.Concat(OrOperators).ToArray()

        If binOperators.Contains(binExpr.NodeType) Then
            VisitExpression(continuation, binExpr.Left, neg)
            If binOperators.Contains(binExpr.Left.NodeType) And binOperators.Contains(binExpr.Right.NodeType) Then
                If AndOperators.Contains(binExpr.NodeType) Then
                    _encodedQuery &= "NQ"
                End If
                If OrOperators.Contains(binExpr.NodeType) Then
                    _encodedQuery &= "^NQ"
                End If
            End If
            VisitExpression(binExpr.NodeType, binExpr.Right, neg)
        Else
            VisitSimpleExpression(continuation, binExpr, neg)
        End If

    End Sub

    Public Sub VisitSimpleExpression(ByVal continuation As Utilities.ContinuationOperator, ByVal binExpr As Expression, ByVal neg As Boolean)
        Dim myEncodedQuery As New EncodedQueryExpression With {
                .ContinuationOperator = continuation,
                .IsNegated = neg,
                .Operator = binExpr.NodeType,
                .EncodedQuery = _encodedQuery,
                .Expression = binExpr
            }

        If myEncodedQuery.HasValue Then
            _encodedQuery = myEncodedQuery.Value
        End If
    End Sub

    Private Sub VisitContainsExpression(ByVal continuation As Utilities.ContinuationOperator, ByVal methodCall As MethodCallExpression, ByVal neg As Boolean)

        Dim myEncodedQuery As New EncodedQueryExpression With {
            .ContinuationOperator = continuation,
            .IsNegated = neg,
            .Operator = Utilities.RepoExpressionType.IN,
            .EncodedQuery = _encodedQuery,
            .Expression = methodCall
        }

        If myEncodedQuery.HasValue Then
            _encodedQuery = myEncodedQuery.Value
        Else
            VisitExpression(continuation, methodCall, neg)
        End If

    End Sub

    Private Sub VisitSimpleMethodCall(ByVal continuation As Utilities.ContinuationOperator, ByVal methodCall As MethodCallExpression, ByVal neg As Boolean)

        Dim myEncodedQuery As New EncodedQueryExpression With {
            .ContinuationOperator = continuation,
            .IsNegated = neg,
            .Operator = Utilities.GetRepoExpressionType(methodCall.Method.Name),
            .EncodedQuery = _encodedQuery,
            .Expression = methodCall
        }

        If myEncodedQuery.HasValue Then
            _encodedQuery = myEncodedQuery.Value
        Else
            VisitExpression(continuation, methodCall, neg)
        End If

    End Sub
End Class