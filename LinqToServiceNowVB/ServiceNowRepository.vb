Imports System.Linq.Expressions
Imports System.Reflection

Public Class ServiceNowRepository(Of TServiceNow_cmdb_ci_ As New, TGetRecords As New, TGetRecordsResponseGetRecordsResult) ' As New)

    Private proxyUser As New TServiceNow_cmdb_ci_()
    Private _filter As New TGetRecords
    Private _encodedQuery As String
    Private _selectQuery As Func(Of Object, TGetRecordsResponseGetRecordsResult)
    Private _groupbyQuery As String

    Protected Friend Property Credential As System.Net.NetworkCredential

    Private Sub SetFilterProperty(prop As String, val As String)

        Dim t As Type = _filter.[GetType]()
        Dim fInfo As FieldInfo = t.GetField(prop)

        If fInfo IsNot Nothing Then
            fInfo.SetValue(_filter, val)
        Else
            Dim pInfo As PropertyInfo = t.GetProperty(prop)

            If prop = "__order_by" Then
                Dim existing_value As Object = pInfo.GetValue(_filter, Nothing)
                If existing_value IsNot Nothing Then
                    val = existing_value.ToString() & "," & val
                End If
            End If

            pInfo.SetValue(_filter, val, Nothing)
        End If

    End Sub

    Private Sub SetOrdering(ByVal order As String, ByVal field As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Object)))
        SetFilterProperty(order, GetOrdering(field))
    End Sub

    Private Function GetOrdering(ByVal field As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Object))) As String

        If field.Body.NodeType = ExpressionType.New Then
            Return String.Join(",", CType(field.Body, NewExpression).Arguments.Select(Function(o) Utilities.GetPropertyName(o)).ToArray())
        Else
            Return Utilities.GetPropertyName(field.Body)
        End If

    End Function

    Private Function IsLimited() As Boolean
        Dim retVal As Boolean

        Dim t As Type = _filter.GetType()

        Dim propInfo As PropertyInfo() = t.GetProperties()

        retVal = propInfo.Any(Function(o) {"__last_row", "__limit"}.Contains(o.Name) And o.GetValue(_filter, Nothing) IsNot Nothing)

        Return retVal
    End Function

    Private Function GetFirstRow() As String
        Dim retVal As String = "0"

        Dim t As Type = _filter.GetType()

        Dim propInfo As PropertyInfo = t.GetProperty("__first_row")

        Dim obj As Object = propInfo.GetValue(_filter, Nothing)

        If obj IsNot Nothing Then
            retVal = obj.ToString()
        End If

        Return retVal
    End Function

    Private Sub SetWebReferenceCredentials(ByVal info As MemberInfo())

        For Each p As PropertyInfo In info.Where(Function(o) o.MemberType = MemberTypes.Property)
            p.SetValue(proxyUser, Credential, Nothing)
        Next

        For Each f As FieldInfo In info.Where(Function(o) o.MemberType = MemberTypes.Field)
            f.SetValue(proxyUser, Credential)
        Next

    End Sub

    Private Sub SetServiceReferenceCredentials(ByVal info As MemberInfo())

        For Each p As PropertyInfo In info.Where(Function(o) o.MemberType = MemberTypes.Property)
            Dim pUserName As PropertyInfo = p.PropertyType.GetProperty("UserName")
            Dim userName As Object = pUserName.GetValue(p.GetValue(proxyUser, Nothing), Nothing)
            pUserName.PropertyType.GetProperty("UserName").SetValue(userName, Credential.UserName, Nothing)
            pUserName.PropertyType.GetProperty("Password").SetValue(userName, Credential.Password, Nothing)
        Next

        For Each f As FieldInfo In info.Where(Function(o) o.MemberType = MemberTypes.Field)
            Dim pUserName As PropertyInfo = f.FieldType.GetProperty("UserName")
            Dim userName As Object = pUserName.GetValue(f.GetValue(proxyUser), Nothing)
            pUserName.PropertyType.GetProperty("UserName").SetValue(userName, Credential.UserName, Nothing)
            pUserName.PropertyType.GetProperty("Password").SetValue(userName, Credential.Password, Nothing)
        Next

    End Sub

    Private Sub SetCredentials(ByVal t As Type)

        If Credential IsNot Nothing Then
            Try
                If t.BaseType Is GetType(System.Web.Services.Protocols.SoapHttpClientProtocol) Then
                    SetWebReferenceCredentials(t.GetMember("Credentials"))
                Else
                    SetServiceReferenceCredentials(t.GetMember("ClientCredentials"))
                End If
            Catch ex As Exception
                Throw New Exception("Exception while setting security credentials for web service.")
            End Try
        End If

    End Sub

    Private Function GetRecords() As TGetRecordsResponseGetRecordsResult()

        Dim t As Type = proxyUser.GetType()

        SetCredentials(t)

        Dim methodInfo As MethodInfo = t.GetMethod("getRecords")

        AppendGroupBy()

        SetFilterProperty("__encoded_query", _encodedQuery)

        If Not IsLimited() Then
            Dim list As New List(Of TGetRecordsResponseGetRecordsResult)
            Dim first As Integer = GetFirstRow()
            Dim last As Integer = 250

            Dim ranged = Range(first, last)
            Dim rslt As Object() = methodInfo.Invoke(proxyUser, {ranged._filter})

            Do Until rslt.Count = 0
                If _selectQuery IsNot Nothing Then
                    list.AddRange(rslt.Select(Function(o) _selectQuery(o)).ToArray())
                Else
                    list.AddRange(rslt)
                End If
                first += 250
                ranged = Range(first, last)
                rslt = methodInfo.Invoke(proxyUser, {ranged._filter})
            Loop

            Return list.ToArray()
        Else
            If _selectQuery IsNot Nothing Then
                Dim ret As Object() = methodInfo.Invoke(proxyUser, {_filter})
                Return ret.Select(Function(o) _selectQuery(o)).ToArray()
            Else
                Dim ret As TGetRecordsResponseGetRecordsResult() = methodInfo.Invoke(proxyUser, {_filter})
                Return ret.ToArray()
            End If
        End If

    End Function

    Private Sub AppendGroupBy()

        If String.IsNullOrEmpty(_groupbyQuery) Then
            Exit Sub
        End If

        _encodedQuery &= _groupbyQuery

    End Sub

    Public Function ToArray() As TGetRecordsResponseGetRecordsResult()

        Dim ret As TGetRecordsResponseGetRecordsResult() = GetRecords()

        Return ret.ToArray()

    End Function

    Public Function ToList() As List(Of TGetRecordsResponseGetRecordsResult)
        Return ToArray().ToList()
    End Function

    Public Function ToDictionary(Of U)(keySelector As Func(Of TGetRecordsResponseGetRecordsResult, U)) As Dictionary(Of U, TGetRecordsResponseGetRecordsResult)

        Dim ret As TGetRecordsResponseGetRecordsResult() = GetRecords()

        Return ret.ToDictionary(Function(o) keySelector(o))

    End Function

    Public Function ToDictionary(Of U, V)(ByVal keySelector As Func(Of TGetRecordsResponseGetRecordsResult, U), ByVal elementSelector As Func(Of TGetRecordsResponseGetRecordsResult, V)) As Dictionary(Of U, V)

        Dim ret As TGetRecordsResponseGetRecordsResult() = GetRecords()

        Dim dict = ret.ToDictionary(Function(o) keySelector(o), Function(o) elementSelector(o))

        Return dict

    End Function

    Public Function [Select](Of U)(ByVal selector As Func(Of TGetRecordsResponseGetRecordsResult, U)) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, U)

        Dim other As New ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, U)()
        other._selectQuery = Function(s) selector(s)
        other.Credential = Me.Credential()
        other._encodedQuery = _encodedQuery
        other._groupbyQuery = _groupbyQuery

        FillCopy(other)

        Return other

    End Function

    Public Function Where(ByVal stmt As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Boolean))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        If Not String.IsNullOrEmpty(_encodedQuery) AndAlso Not _encodedQuery.EndsWith("NQ") Then
            _encodedQuery &= "NQ"
        End If

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal._encodedQuery &= (New ExpressionVisitor).VisitExpression(Utilities.ContinuationOperator.And, stmt.Body)

        Return retVal

    End Function

    Public Function SkipWhile(ByVal stmt As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Boolean))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal._encodedQuery &= "^" & (New ExpressionVisitor).VisitExpression(Utilities.ContinuationOperator.And, stmt.Body, True)

        Return retVal

    End Function

    Public Function TakeWhile(ByVal stmt As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Boolean))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal._encodedQuery &= "^" & (New ExpressionVisitor).VisitExpression(Utilities.ContinuationOperator.And, stmt.Body)

        Return retVal

    End Function

    Public Function GroupBy(Of U, V)(keySelector As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, U)),
                                     elementSelector As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, V))) _
                                 As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, V)
        Dim selected = [Select](elementSelector.Compile())
        selected._groupbyQuery = GroupBy(keySelector)._groupbyQuery
        Return selected
    End Function

    Public Function GroupBy(Of U, V, W)(keySelector As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, U)), _
                                        elementSelector As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, V)), _
                                        resultSelector As Expression(Of Func(Of U, IEnumerable(Of TGetRecordsResponseGetRecordsResult), W))) _
                                    As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, V)
        Dim selected = [Select](elementSelector.Compile())
        selected._groupbyQuery = GroupBy(keySelector)._groupbyQuery
        Return selected
    End Function

    Public Function GroupBy(Of U, V)(keySelector As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, U)),
                                     resultSelector As Expression(Of Func(Of U, IEnumerable(Of TGetRecordsResponseGetRecordsResult), V))) _
                                 As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)
        Return GroupBy(keySelector)
    End Function

    Public Function GroupBy(Of U)(ByVal field As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, U))) _
        As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim setGroupBy = Sub(expr As Expression)
                             Dim query As String

                             If expr.NodeType = ExpressionType.New Then
                                 query = "^GROUPBY" & String.Join(",", CType(expr, NewExpression).Arguments.Select(Function(o) Utilities.GetPropertyName(o)).ToArray())
                             Else
                                 query = "^GROUPBY" & Utilities.GetPropertyName(expr)
                             End If

                             _groupbyQuery &= query
                         End Sub

        If field.Body.NodeType = ExpressionType.Convert Then
            Dim unaryExpression As UnaryExpression = CType(field.Body, UnaryExpression)
            setGroupBy(unaryExpression.Operand)
        Else
            setGroupBy(field.Body)
        End If

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        Return retVal

    End Function

    Public Function Join(Of T As New, U As New, V As New)(serviceNowRepository As ServiceNowRepository(Of T, U, V),
                  outerKeySelector As Func(Of TGetRecordsResponseGetRecordsResult, Object),
                  innerKeySelector As Func(Of V, Object),
                  resultSelector As Func(Of TGetRecordsResponseGetRecordsResult, V, Object)) As IEnumerable(Of Object)
        Dim query = Me.ToList().Join(serviceNowRepository.ToList(),
                                Function(o) outerKeySelector(o),
                                Function(o) innerKeySelector(o),
                                Function(o, p) resultSelector(o, p))

        Return query
    End Function

    Public Function Join(Of T As New, U As New, V As New, W)(serviceNowRepository As ServiceNowRepository(Of T, U, V),
                  outerKeySelector As Func(Of TGetRecordsResponseGetRecordsResult, Object),
                  innerKeySelector As Func(Of V, Object),
                  resultSelector As Func(Of TGetRecordsResponseGetRecordsResult, V, W)) As IEnumerable(Of Object)
        Dim query = Me.ToList().Join(serviceNowRepository.ToList(),
                                Function(o) outerKeySelector(o),
                                Function(o) innerKeySelector(o),
                                Function(o, p) resultSelector(o, p))

        Return query
    End Function

    Public Function First(ByVal stmt As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Boolean))) As Object

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal.SetFilterProperty("__limit", "1")

        If stmt.Body.NodeType <> ExpressionType.Constant Then
            retVal.Where(stmt)
        End If

        Return retVal.ToArray().FirstOrDefault()

    End Function

    Public Function First() As Object
        Return First(Function(o) True)
    End Function

    Public Function [Single](ByVal stmt As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Boolean))) As Object
        Return First(stmt)
    End Function

    Public Function [Single]() As Object
        Return First()
    End Function

    Public Function Any(ByVal stmt As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Boolean))) As Boolean
        Return First(stmt) IsNot Nothing
    End Function

    Public Function Any() As Boolean
        Return Any(Function(o) True)
    End Function

    Public Function Take(ByVal count As Integer) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        Dim lastRow As Integer = count + Integer.Parse(GetFirstRow())

        retVal.SetFilterProperty("__last_row", lastRow)

        Return retVal

    End Function

    Private Function Range(ByVal start As Integer, ByVal last As Integer) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal.SetFilterProperty("__first_row", start.ToString())
        retVal.SetFilterProperty("__last_row", (start + last).ToString())

        Return retVal

    End Function

    Public Function ElementAt(ByVal at As Integer) As Object

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        Return retVal.Range(at - 1, at).First()

    End Function

    Public Function DeepCopy() As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim other As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) = DirectCast(Me.MemberwiseClone(), ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult))

        FillCopy(other)

        Return other

    End Function

    Private Sub FillCopy(Of TGetRecs As New, TGetRecsResponseGetRecsResult)(ByVal other As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecs, TGetRecsResponseGetRecsResult))
        other._filter = New TGetRecs()

        Dim fieldNames = {"__encoded_query", "__limit", "__first_row", "__last_row", "__order_by", "__order_by_desc"}

        Dim t As Type = _filter.[GetType]()
        t.GetFields().Where(Function(o) fieldNames.Contains(o.Name) AndAlso o.GetValue(_filter) IsNot Nothing).ToList().
            ForEach(Sub(o) other.SetFilterProperty(o.Name, o.GetValue(_filter)))

        t.GetProperties().Where(Function(o) fieldNames.Contains(o.Name) AndAlso o.GetValue(_filter, Nothing) IsNot Nothing).ToList().
            ForEach(Sub(o) other.SetFilterProperty(o.Name, o.GetValue(_filter, Nothing)))
    End Sub

    Public Function Skip(ByVal count As Integer) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal.SetFilterProperty("__first_row", count.ToString())

        Return retVal

    End Function

    Public Function OrderBy(ByVal field As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Object))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal.SetOrdering("__order_by", field)

        Return retVal

    End Function

    Public Function ThenBy(source As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Object))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)
        Return OrderBy(source)
    End Function

    Public Function OrderByDescending(ByVal field As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Object))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)

        Dim retVal As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult) =
            Me.DeepCopy()

        retVal.SetOrdering("__order_by_desc", field)

        Return retVal

    End Function

    Public Function ThenByDescending(source As Expression(Of Func(Of TGetRecordsResponseGetRecordsResult, Object))) As ServiceNowRepository(Of TServiceNow_cmdb_ci_, TGetRecords, TGetRecordsResponseGetRecordsResult)
        Return OrderByDescending(source)
    End Function

    Public Function Insert(Of TInsert)(ByVal _insert As TInsert) As String

        Dim t As Type = proxyUser.GetType()

        SetCredentials(t)

        Dim methodInfo As MethodInfo = t.GetMethod("insert")

        Return methodInfo.Invoke(proxyUser, {_insert}).sys_id

    End Function

    Public Function Update(Of TUpdate)(ByVal _update As TUpdate) As String

        Dim t As Type = proxyUser.GetType()

        SetCredentials(t)

        Dim methodInfo As MethodInfo = t.GetMethod("update")

        Return methodInfo.Invoke(proxyUser, {_update}).sys_id

    End Function

    Public Function Delete(Of TDelete As New)(ByVal _delete As TDelete) As String

        Dim t As Type = proxyUser.GetType()

        SetCredentials(t)

        Dim methodInfo As MethodInfo = t.GetMethod("deleteRecord")

        Return methodInfo.Invoke(proxyUser, {_delete}).count

    End Function

    Public Function GetEnumerator() As IEnumerator(Of TGetRecordsResponseGetRecordsResult)
        Return ToList().GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator
        Return Me.GetEnumerator()
    End Function

    Public Sub New()

    End Sub

    Public Sub New(ByVal credential As System.Net.NetworkCredential)
        Me.Credential = credential
    End Sub

End Class