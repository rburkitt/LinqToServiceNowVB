Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Reflection

Public Class Utilities
    Public Enum RepoExpressionType
        <Description("=")>
        Equal = ExpressionType.Equal
        <Description("!=")>
        NotEqual = ExpressionType.NotEqual
        <Description(">")>
        GreaterThan = ExpressionType.GreaterThan
        <Description("<")>
        LessThan = ExpressionType.LessThan
        <Description(">=")>
        GreaterThanOrEqual = ExpressionType.GreaterThanOrEqual
        <Description("<=")>
        LessThanOrEqual = ExpressionType.LessThanOrEqual
        [IN] = 100
        [STARTSWITH] = 200
        [CONTAINS] = 300
        [ENDSWITH] = 400
        <Description("LIKE")>
        LIKESTRING = 500
    End Enum

    Public Enum ContinuationOperator
        <Description("^")>
        [And] = ExpressionType.And
        <Description("^")>
        [AndAlso] = ExpressionType.AndAlso
        <Description("^OR")>
        [Or] = ExpressionType.Or
        <Description("^OR")>
        [OrElse] = ExpressionType.OrElse
    End Enum

    Private Shared Function GetOperator(Of T)(ByVal en As T) As String
        Dim type As Type = en.GetType()

        Dim memInfo As MemberInfo() = type.GetMember(en.ToString())

        If memInfo IsNot Nothing And memInfo.Length > 0 Then
            Dim attrs As Object() = memInfo(0).GetCustomAttributes(GetType(DescriptionAttribute), False)

            If attrs IsNot Nothing And attrs.Length > 0 Then
                Return CType(attrs(0), DescriptionAttribute).Description
            End If
        End If

        Return en.ToString()
    End Function

    Public Shared Function GetContinuationOperator(ByVal en As ContinuationOperator) As String
        Return GetOperator(Of ContinuationOperator)(en)
    End Function

    Public Shared Function GetRepoExpressionType(ByVal en As RepoExpressionType) As String
        Return GetOperator(Of RepoExpressionType)(en)
    End Function

    Public Shared Function GetRepoExpressionType(ByVal en As String) As String
        Dim val = DirectCast([Enum].Parse(GetType(RepoExpressionType), en.ToUpper()), RepoExpressionType)
        Return val
    End Function

    Public Shared Function NegateRepoExpressionType(ByVal en As RepoExpressionType) As String

        Dim oper As String = GetRepoExpressionType(en)

        Select Case en
            Case RepoExpressionType.IN
                oper = "NOT IN"
            Case RepoExpressionType.CONTAINS
                oper = "DOES NOT CONTAIN"
            Case RepoExpressionType.STARTSWITH
                oper = "NOT STARTSWITH"
            Case RepoExpressionType.ENDSWITH
                oper = "NOT ENDSWITH"
            Case Else
                oper = GetRepoExpressionType(FlipRepoExpressionType(en))
        End Select

        Return oper

    End Function

    Public Shared Function FlipRepoExpressionType(ByVal en As RepoExpressionType) As RepoExpressionType

        Dim oper As RepoExpressionType

        Select Case en
            Case RepoExpressionType.Equal
                oper = RepoExpressionType.Equal
            Case ExpressionType.GreaterThan
                oper = ExpressionType.LessThanOrEqual
            Case ExpressionType.GreaterThanOrEqual
                oper = ExpressionType.LessThan
            Case ExpressionType.LessThan
                oper = ExpressionType.GreaterThanOrEqual
            Case ExpressionType.LessThanOrEqual
                oper = ExpressionType.GreaterThan
        End Select

        Return oper

    End Function

    Public Shared Function GetPropertyName(propertyRefExpr As Expression) As String

        If propertyRefExpr Is Nothing Then
            Throw New ArgumentNullException("propertyRefExpr", "propertyRefExpr is null.")
        End If

        If propertyRefExpr.NodeType = ExpressionType.Constant Then
            Return propertyRefExpr.ToString()
        End If

        Dim memberExpr As MemberExpression = TryCast(propertyRefExpr, MemberExpression)
        If memberExpr Is Nothing Then
            Dim unaryExpr As UnaryExpression = TryCast(propertyRefExpr, UnaryExpression)
            If unaryExpr IsNot Nothing AndAlso unaryExpr.NodeType = ExpressionType.Convert Then
                memberExpr = TryCast(unaryExpr.Operand, MemberExpression)
            End If
        End If

        If memberExpr IsNot Nothing AndAlso memberExpr.Member.MemberType = System.Reflection.MemberTypes.[Property] Then
            Return memberExpr.Member.Name
        End If

        If memberExpr IsNot Nothing AndAlso memberExpr.Member.MemberType = System.Reflection.MemberTypes.Field Then
            Return memberExpr.Member.Name
        End If

        Throw New ArgumentException("No property reference expression was found.", "propertyRefExpr")

    End Function
End Class
