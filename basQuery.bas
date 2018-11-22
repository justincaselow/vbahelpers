Attribute VB_Name = "basQuery"
Function GetIndexFromDelimited(ByVal StringValue As String, ByVal delimiter As String, ByVal index As Integer)
    
    result = Split(StringValue, delimiter)
    
    GetIndexFromDelimited = result(index)

End Function

Function IsInArray(valToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, valToBeFound)) > -1)
End Function
