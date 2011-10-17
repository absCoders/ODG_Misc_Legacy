Public Module UtilityFunctions
    Public Function IsNullableType(Of T)(ByVal myObj As T) As Boolean
        If myObj Is Nothing Then Return True
        Return GetType(T).IsGenericType AndAlso GetType(T).GetGenericTypeDefinition().Equals(GetType(Nullable(Of )))
    End Function

    Public Function GetObjType(Of T)(ByVal myObj As T) As Type
        If IsNullableType(myObj) Then
            Return Nullable.GetUnderlyingType(GetType(T))
        Else
            Return myObj.GetType()
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Left(ByVal str As String, ByVal length As Integer)
        If str Is Nothing Then Return ""
        Return str.Substring(0, Math.Min(str.Length, length))
    End Function

End Module