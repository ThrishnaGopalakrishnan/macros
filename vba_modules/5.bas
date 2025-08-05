Attribute VB_Name = "Module5"
Public Sub component()
C_changetype = Worksheets("Attribute").Cells(3, 3).Value
C_subdatatype = Worksheets("subattribute").Cells(5, 2).Value
C_enumeration = Worksheets("Attribute").Cells(16, 3).Value

If C_subdatatype = "SymbolStruct@1.0" Or C_subdatatype = "BoolStruct@1.0" Then
        Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
        Sheets("constraint").Cells(2, 2).Value = "*"
        Sheets("constraint").Cells(2, 3).Value = "10480"
        Sheets("constraint").Cells(2, 4).Value = "contributor"
        Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(2, 6).Value = "max_occurs"
        Sheets("constraint").Cells(2, 8).Value = 1
        Sheets("constraint").Cells(3, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(3, 2).Value = "*"
        Sheets("constraint").Cells(3, 3).Value = "10480"
        Sheets("constraint").Cells(3, 4).Value = "contributor"
        Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(3, 6).Value = "max_occurs"
        Sheets("constraint").Cells(3, 8).Value = 1
        Sheets("constraint").Cells(4, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(4, 2).Value = "value"
        Sheets("constraint").Cells(4, 3).Value = "10480"
        Sheets("constraint").Cells(4, 4).Value = "contributor"
        Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(4, 6).Value = "enumeration"
    End If

    If C_subdatatype = "StringStruct@1.0" Then
        Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
        Sheets("constraint").Cells(2, 2).Value = "*"
        Sheets("constraint").Cells(2, 3).Value = "10480"
        Sheets("constraint").Cells(2, 4).Value = "contributor"
        Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(2, 6).Value = "max_occurs"
        Sheets("constraint").Cells(2, 8).Value = 1
        Sheets("constraint").Cells(3, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(3, 2).Value = "*"
        Sheets("constraint").Cells(3, 3).Value = "10480"
        Sheets("constraint").Cells(3, 4).Value = "contributor"
        Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(3, 6).Value = "max_occurs"
        Sheets("constraint").Cells(3, 8).Value = 1
        Sheets("constraint").Cells(4, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(4, 2).Value = "value"
        Sheets("constraint").Cells(4, 3).Value = "10480"
        Sheets("constraint").Cells(4, 4).Value = "contributor"
        Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(4, 6).Value = "max_length"
          
    End If


    If C_subdatatype = "LocalizedString@1.0" Then
        Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
        Sheets("constraint").Cells(2, 2).Value = "*"
        Sheets("constraint").Cells(2, 3).Value = "10480"
        Sheets("constraint").Cells(2, 4).Value = "contributor"
        Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(2, 6).Value = "max_occurs"
        Sheets("constraint").Cells(2, 8).Value = 1
        Sheets("constraint").Cells(3, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(3, 2).Value = "*"
        Sheets("constraint").Cells(3, 3).Value = "10480"
        Sheets("constraint").Cells(3, 4).Value = "contributor"
        Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(3, 6).Value = "max_occurs"
        Sheets("constraint").Cells(3, 8).Value = 1
        Sheets("constraint").Cells(4, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(4, 2).Value = "value"
        Sheets("constraint").Cells(4, 3).Value = "10480"
        Sheets("constraint").Cells(4, 4).Value = "contributor"
        Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(4, 6).Value = "max_length"
          
    End If

If C_subdatatype = "UnitDecimalStruct@1.0" Then
        Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
        Sheets("constraint").Cells(2, 2).Value = "*"
        Sheets("constraint").Cells(2, 3).Value = "10480"
        Sheets("constraint").Cells(2, 4).Value = "contributor"
        Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(2, 6).Value = "max_occurs"
        Sheets("constraint").Cells(2, 8).Value = 1
        Sheets("constraint").Cells(3, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(3, 2).Value = "*"
        Sheets("constraint").Cells(3, 3).Value = "10480"
        Sheets("constraint").Cells(3, 4).Value = "contributor"
        Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(3, 6).Value = "max_occurs"
        Sheets("constraint").Cells(3, 8).Value = 1
        Sheets("constraint").Cells(4, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(4, 2).Value = "value"
        Sheets("constraint").Cells(4, 3).Value = "10480"
        Sheets("constraint").Cells(4, 4).Value = "contributor"
        Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(4, 6).Value = "min_value"
        Sheets("constraint").Cells(5, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(5, 2).Value = "value"
        Sheets("constraint").Cells(5, 3).Value = "10480"
        Sheets("constraint").Cells(5, 4).Value = "contributor"
        Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(5, 6).Value = "max_value"
        Sheets("constraint").Cells(6, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(6, 2).Value = "value"
        Sheets("constraint").Cells(6, 3).Value = "10480"
        Sheets("constraint").Cells(6, 4).Value = "contributor"
        Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(6, 6).Value = "max_scale"
        Sheets("constraint").Cells(7, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(7, 2).Value = "value"
        Sheets("constraint").Cells(7, 3).Value = "10480"
        Sheets("constraint").Cells(7, 4).Value = "contributor"
        Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(7, 6).Value = "max_precision"
        Sheets("constraint").Cells(7, 8).Value = 12
        Sheets("constraint").Cells(8, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(8, 2).Value = "unit"
        Sheets("constraint").Cells(8, 3).Value = "10480"
        Sheets("constraint").Cells(8, 4).Value = "contributor"
        Sheets("constraint").Cells(8, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(8, 6).Value = "enumeration"
    End If
    
    If C_subdatatype = "DecimalStruct@1.0" Or C_subdatatype = "Price@1.0" Then
        Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
        Sheets("constraint").Cells(2, 2).Value = "*"
        Sheets("constraint").Cells(2, 3).Value = "10480"
        Sheets("constraint").Cells(2, 4).Value = "contributor"
        Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(2, 6).Value = "max_occurs"
        Sheets("constraint").Cells(2, 8).Value = 1
        Sheets("constraint").Cells(3, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(3, 2).Value = "value"
        Sheets("constraint").Cells(3, 3).Value = "10480"
        Sheets("constraint").Cells(3, 4).Value = "contributor"
        Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(3, 6).Value = "max_value"
        Sheets("constraint").Cells(4, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(4, 2).Value = "value"
        Sheets("constraint").Cells(4, 3).Value = "10480"
        Sheets("constraint").Cells(4, 4).Value = "contributor"
        Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(4, 6).Value = "min_value"
        Sheets("constraint").Cells(5, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(5, 2).Value = "value"
        Sheets("constraint").Cells(5, 3).Value = "10480"
        Sheets("constraint").Cells(5, 4).Value = "contributor"
        Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(5, 6).Value = "max_scale"
        Sheets("constraint").Cells(6, 1).Value = Worksheets("subattribute").Cells(3, 2).Value
        Sheets("constraint").Cells(6, 2).Value = "value"
        Sheets("constraint").Cells(6, 3).Value = "10480"
        Sheets("constraint").Cells(6, 4).Value = "contributor"
        Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
        Sheets("constraint").Cells(6, 6).Value = "max_precision"
        Sheets("constraint").Cells(6, 8).Value = 12
 End If

End Sub




