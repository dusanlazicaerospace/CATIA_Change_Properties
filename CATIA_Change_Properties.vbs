Sub CATMain()
    
    Dim product 
    Set product = CATIA.ActiveDocument.Product
    
    
    ChangeParameterInProduct product
End Sub

Sub ChangeParameterInProduct
    Dim subProduct As Product
    Dim part As Part
    
    
    ChangeParameterInCurrentProduct product
    
    
    For Each subProduct In product.Products
        ChangeParameterInProduct subProduct
    Next 
    
    For Each part In product.Products
        ChangeParameterInPart part
    Next 
End Sub

Sub ChangeParameterInCurrentProduct
    If ParameterExists(product, "important") Then
        product.Parameters.Item("important").Value = "Yes"
    End If
End Sub

Sub ChangeParameterInPart
    If ParameterExists(part, "important") Then
        part.Parameters.Item("important").Value = "Yes"
    End If
End Sub

Function ParameterExists
    On Error Resume Next
    ParameterExists = Not (object.Parameters.Item(parameterName) Is Nothing)
    On Error GoTo 0
End Function