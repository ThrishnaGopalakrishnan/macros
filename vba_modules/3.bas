Attribute VB_Name = "Module3"
Sub Button10_Click()
C_datatype = Worksheets("Attribute").Cells(15, 3).Value
C_changetype = Worksheets("Attribute").Cells(4, 3).Value
C_datatypestat = Worksheets("Attribute").Cells(14, 5).Value
C_subdatatype = Worksheets("subattribute").Cells(5, 2).Value
C_concept_name = Worksheets("Attribute").Cells(11, 3).Value
C_enumstat = Worksheets("Attribute").Cells(17, 5).Value
C_Enumname = Worksheets("Attribute").Cells(17, 3).Value
C_subenum_name = Worksheets("subattribute").Cells(7, 2).Value
C_Descriptive = Worksheets("Attribute").Cells(31, 3).Value
C_Identifier = Worksheets("Attribute").Cells(32, 3).Value
C_Branded = Worksheets("Attribute").Cells(33, 3).Value
C_Identity = Worksheets("Attribute").Cells(30, 3).Value
C_Internal = Worksheets("Attribute").Cells(28, 3).Value
c_listing_group = Worksheets("Attribute").Cells(28, 4).Value
C_configurable = Worksheets("Attribute").Cells(35, 3).Value
C_refinement = Worksheets("Attribute").Cells(34, 3).Value
C_Context = Worksheets("Attribute").Cells(19, 3).Value
C_replacement = Worksheets("attribute").Cells(36, 3).Value
C_replaced = Worksheets("attribute").Cells(36, 4).Value
C_subconcept = Worksheets("contype").Cells(22, 18).Value
C_subdescriptive = Worksheets("subattribute").Cells(21, 2).Value
C_subidentifier = Worksheets("subattribute").Cells(22, 2).Value
C_subbranded = Worksheets("subattribute").Cells(23, 2).Value
C_subrefinement = Worksheets("subattribute").Cells(24, 2).Value
C_subreplacement = Worksheets("subattribute").Cells(26, 2).Value
C_subreplaced = Worksheets("subattribute").Cells(26, 3).Value
C_subenumtype = Worksheets("subattribute").Cells(7, 3).Value
C_contains_maxprecision = Worksheets("contype").Cells(2, 18).Value
C_contains_scale = Worksheets("contype").Cells(3, 18).Value
C_contains_unit = Worksheets("contype").Cells(4, 18).Value
C_contains_provide = Worksheets("contype").Cells(5, 18).Value
C_contains_enum = Worksheets("contype").Cells(6, 18).Value
C_contains_minval = Worksheets("contype").Cells(7, 18).Value
C_contains_maxval = Worksheets("contype").Cells(8, 18).Value
C_contains_maxlength = Worksheets("contype").Cells(9, 18).Value
C_contains_underscore = Worksheets("contype").Cells(10, 18).Value
C_contains_space = Worksheets("contype").Cells(11, 18).Value
C_visible = Worksheets("contype").Cells(12, 18).Value
C_enum_contains_space = Worksheets("contype").Cells(13, 18).Value
C_dis_contains_unit = Worksheets("contype").Cells(14, 18).Value
C_contains_eligible = Worksheets("contype").Cells(15, 18).Value
C_contains_partially_eligible = Worksheets("contype").Cells(16, 18).Value
C_contains_ineligible = Worksheets("contype").Cells(17, 18).Value
C_contains_datatype = Worksheets("Display Properties").Cells(17, 1).Value
C_subtype = Worksheets("Attribute").Cells(37, 3).Value
C_isenum = Worksheets("contype").Cells(19, 18).Value
C_hassubtypecls = Worksheets("contype").Cells(24, 18).Value
C_subisenum = Worksheets("contype").Cells(24, 18).Value
C_subenum_space = Worksheets("contype").Cells(16, 18).Value
C_ac1 = Worksheets("cust rel requirements").Cells(2, 5).Value
C_rk1 = Worksheets("cust rel requirements").Cells(2, 8).Value
C_acp1 = Worksheets("cust rel requirements").Cells(2, 3).Value
C_acd1 = Worksheets("cust rel requirements").Cells(2, 2).Value
C_ac2 = Worksheets("cust rel requirements").Cells(3, 5).Value
C_rk2 = Worksheets("cust rel requirements").Cells(3, 8).Value
C_acp2 = Worksheets("cust rel requirements").Cells(3, 3).Value
C_acd2 = Worksheets("cust rel requirements").Cells(3, 2).Value
C_ac3 = Worksheets("cust rel requirements").Cells(4, 5).Value
C_rk3 = Worksheets("cust rel requirements").Cells(4, 8).Value
C_acp3 = Worksheets("cust rel requirements").Cells(4, 3).Value
C_acd3 = Worksheets("cust rel requirements").Cells(4, 2).Value
C_ac4 = Worksheets("cust rel requirements").Cells(5, 5).Value
C_rk4 = Worksheets("cust rel requirements").Cells(5, 8).Value
C_acp4 = Worksheets("cust rel requirements").Cells(5, 3).Value
C_acd4 = Worksheets("cust rel requirements").Cells(5, 2).Value
C_ac5 = Worksheets("cust rel requirements").Cells(6, 5).Value
C_rk5 = Worksheets("cust rel requirements").Cells(6, 8).Value
C_acp5 = Worksheets("cust rel requirements").Cells(6, 3).Value
C_acd5 = Worksheets("cust rel requirements").Cells(6, 2).Value

Dim iRow As Long, lastRow As Long, firstRow As Long, ARow As Long, firstARow As Long, lastARow As Long

Application.Calculation = xlManual
Application.DisplayStatusBar = False
Application.ScreenUpdating = False


'Clear errors sheet (so that it can be rerun after errors corrected)

Sheets("Errors").Rows("2:" & Rows.Count).EntireRow.Delete



'Define first/last rows
With Sheets("Attribute").Select

    lastRow = Sheets("Attribute").Range("A" & Sheets("Attribute").Rows.Count).End(xlUp).Row
    firstRow = Sheets("Attribute").Range("1:37").Row
    lastARow = 2
    
'Clear Previous Errors
    Sheets("Attribute").Cells(17, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(3, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(17, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Display Properties").Cells(17, 1).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(19, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(37, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(15, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(17, 4).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(30, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(35, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(29, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(31, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(28, 4).Interior.Color = RGB(255, 255, 255)
    Sheets("Display Properties").Cells(5, 4).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(20, 2).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(21, 2).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(20, 2).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(7, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(7, 2).Interior.Color = RGB(255, 255, 255)
    Sheets("Attribute").Cells(36, 4).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(25, 3).Interior.Color = RGB(255, 255, 255)
    Sheets("subattribute").Cells(3, 2).Interior.Color = RGB(255, 255, 255)
    
'For each row run validations in columns, if there is an issue fill cell red and write error to errors sheet
    For iRow = firstRow To lastRow
    
              
           If C_concept_name = "*_*" Then
                Sheets("Attribute").Cells(11, 3).Interior.Color = RGB(255, 0, 0)
                Sheets("Errors").Cells(lastARow, 1).Value = "Concept name should not have underscore _"
                lastARow = lastARow + 1
                Else
           End If
         
           
           If C_datatypestat = "New" And C_contains_datatype = "" Then
                  Sheets("Errors").Cells(lastARow, 1).Value = "New Datatype needs to be created on Display Properties Tab"
                  lastARow = lastARow + 1
                  Sheets("Display Properties").Cells(17, 1).Interior.Color = RGB(255, 0, 0)
            Else
           End If
           
           If C_Identity = True And C_Context <> "key_type" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If Identity attribute, then V2 Context Name must be key_type"
                Sheets("Attribute").Cells(19, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
           End If
                    
           
           If C_contains_underscore = True Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Display labels should not have underscores"
                Sheets("Attribute").Cells(37, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
           End If
           
            If C_enumstat = "See Enum Tab" Then
                Sheets("Attribute").Cells(17, 4).Value = "Not Applicable"
                Sheets("Attribute").Cells(17, 3).Value = "Not Applicable"
                Else
           End If
                   
           
           If C_contains_space > 0 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Variant in constraints tabs contains a space. Proper variant format would be x_x if more than one word."
                lastARow = lastARow + 1
           End If
           
           If C_changetype = "Component Attribute" Or C_changetype = "New Complex Attribute" Or C_changetype = "Complex Attribute With Sub-attribute" Then
                If C_datatype <> "Complex Data Type" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Datatype for Complex or Base Attribute for Component should be Complex Data Type."
                Sheets("Attribute").Cells(15, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
                End If
           End If
           
           If C_contains_provide = "false" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Short Descriptions must start with the word Provide."
                Sheets("Display Properties").Cells(5, 4).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
           End If
           
           If Sheets("Attribute").Cells(3, 3) = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Link to SIM must be provided."
                Sheets("Attribute").Cells(3, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
           End If
         
             
           If C_ac2 = "customer_relevant_attribute" And C_rk2 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac3 = "customer_relevant_attribute" And C_rk3 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac4 = "customer_relevant_attribute" And C_rk4 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac5 = "customer_relevant_attribute" And C_rk5 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
            If C_ac1 = "customer_relevant_attribute" And C_acp1 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Product Type is required for Customer Relevant Attribute."
                lastARow = lastARow + 1
                Else
           End If
           
            If C_ac2 = "customer_relevant_attribute" And C_acp2 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac3 = "customer_relevant_attribute" And C_acp3 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac4 = "customer_relevant_attribute" And C_acp4 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac5 = "customer_relevant_attribute" And C_acp5 = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Ranking must be added for each customer relevant attribute."
                lastARow = lastARow + 1
                Else
           End If
        
                 
           If C_ac2 = "customer_relevant_attribute" And C_acd2 <> "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Domain should not be included for Customer Relevant Attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac3 = "customer_relevant_attribute" And C_acd3 <> "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Domain should not be included for Customer Relevant Attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac4 = "customer_relevant_attribute" And C_acd4 <> "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Domain should not be included for Customer Relevant Attribute."
                lastARow = lastARow + 1
                Else
           End If
           
           If C_ac5 = "customer_relevant_attribute" And C_acd5 <> "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Domain should not be included for Customer Relevant Attribute."
                lastARow = lastARow + 1
                Else
           End If
           
              
                                 
           If C_changetype = "New Simple Attribute" Then
           
                                                          
                If C_datatype = "SymbolVariant@1.0" And C_Enumname = "" Then
                    Sheets("Attribute").Cells(17, 3).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "Enumeration Name (C16)must be populated for symbol variant."
                    lastARow = lastARow + 1
                    Else
                    
                End If
                              
                If C_replacement = True And C_replaced = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If this is a replacement attribute, fill attribute to be replaced in cell D34."
                Sheets("Attribute").Cells(36, 4).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
                End If
                
                If C_refinement = True Then
                    If C_contains_eligible = 0 Or C_contains_eligible = 0 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If Is Refinement Candidate = True then refinement eligible or refinement partially eligible attribute class must be added on Cust Rel Requirements Tab."
                lastARow = lastARow + 1
                Else
                End If
                End If
                
                If C_refinement = False Then
                    If C_contains_ineligible = 0 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If Is Refinement Candidate = False then refinement ineligible attribute class must be added on Cust Rel Requirements Tab."
                lastARow = lastARow + 1
                Else
                End If
                End If
                               
                                              
                If C_datatype = "SymbolVariant@1.0" And C_Enumname = "Not Applicable" Then
                    Sheets("Attribute").Cells(17, 3).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "Not Applicable is not a valid enumeration name for a symbol variant."
                    lastARow = lastARow + 1
                Else
                End If
                
                If C_datatype = "LocalizedStringVariant@1.0" And C_enumstat = "Closed" Then
                    Sheets("Attribute").Cells(17, 4).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "LocalizedStringVariants cannot be closed set enumerated."
                    lastARow = lastARow + 1
                Else
                End If
                                    
                End If
                       
                If C_datatype = "SymbolVariant@1.0" And C_enumstat <> "Closed" Then
                    Sheets("Attribute").Cells(17, 4).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "Enumeration for symbol variant attribute must be closed at PRODUCT."
                    lastARow = lastARow + 1
                Else
                    
                End If
                       
                If C_contains_unit = 0 Then
                    If C_datatype = "UnitDecimalVariant@1.0" Or C_datatype = "NormalizedUnitDecimalVariant@1.0" Or C_datatype = "UnitIntVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "An enumeration constraint must exist for a Unit Decimal Datatype."
                    lastARow = lastARow + 1
                    End If
                End If
           
                If C_contains_unit <> 2 And C_datatype = "UnitDecimalRangeVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "For a Unit Decimal Range, an enumeration constraint must exist for both value and normalized value."
                    lastARow = lastARow + 1
                End If
                                
                If C_contains_scale = 0 Then
                    If C_datatype = "UnitDecimalVariant@1.0" Or C_datatype = "NormalizedUnitDecimalVariant@1.0" Or C_datatype = "DecimalVariant@1.0" Or C_datatype = "PriceVariant@1.0" Or C_datatype = "FractionVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_scale constraint must exist for a decimal based datatype."
                    lastARow = lastARow + 1
                    End If
                End If
           
                If C_contains_maxprecision = 0 Then
                    If C_datatype = "UnitDecimalVariant@1.0" Or C_datatype = "NormalizedUnitDecimalVariant@1.0" Or C_datatype = "DecimalVariant@1.0" Or C_datatype = "PriceVariant@1.0" Or C_datatype = "FractionVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_precision constraint must exist for a decimal based datatype."
                    lastARow = lastARow + 1
                    End If
                End If
           
                If C_contains_minval = 0 Then
                    If C_datatype = "UnitDecimalVariant@1.0" Or C_datatype = "NormalizedUnitDecimalVariant@1.0" Or C_datatype = "DecimalVariant@1.0" Or C_datatype = "PriceVariant@1.0" Or C_datatype = "FractionVariant@1.0" Or C_datatype = "IntVariant@1.0" Or C_datatype = "TypedIntVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A min_value constraint must exist for a decimal based datatype."
                    lastARow = lastARow + 1
                    End If
                End If
           
                If C_contains_maxval = 0 Then
                    If C_datatype = "UnitDecimalVariant@1.0" Or C_datatype = "NormalizedUnitDecimalVariant@1.0" Or C_datatype = "DecimalVariant@1.0" Or C_datatype = "PriceVariant@1.0" Or C_datatype = "FractionVariant@1.0" Or C_datatype = "IntVariant@1.0" Or C_datatype = "TypedIntVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_value constraint must exist for a decimal based datatype."
                    lastARow = lastARow + 1
                    End If
                End If
                      
                If C_contains_scale <> 2 And C_datatype = "UnitDecimalRangeVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_scale constraint must exist for both the min and max value."
                    lastARow = lastARow + 1
                End If
           
                If C_contains_precision <> 2 And C_datatype = "UnitDecimalRangeVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_Precision constraint must exist for both the min and max value."
                    lastARow = lastARow + 1
                End If
            
                If C_contains_minval <> 2 And C_datatype = "UnitDecimalRangeVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A min_value constraint must exist for both the min and max value."
                    lastARow = lastARow + 1
                End If
           
                If C_contains_maxval <> 2 And C_datatype = "UnitDecimalRangeVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_value constraint must exist for both the min and max value."
                    lastARow = lastARow + 1
                End If
            
                If C_contains_maxlength = 0 Then
                If C_datatype = "LocalizedStringVariant@1.0" Or C_datatype = "StringVariant@1.0" Or C_datatype = "TypedStringVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "A max_length constraint must exist for a string datatype."
                    lastARow = lastARow + 1
                End If
                End If
                                
                If C_dis_contains_unit = 0 Then
                    If C_datatype = "UnitDecimalVariant@1.0" Or C_datatype = "UnitIntVariant@1.0" Or C_datatype = "NormalizedUnitDecimalVariant@1.0" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "Display properties (short_desc, example, label) must exist for the unit if datatype has unit."
                    lastARow = lastARow + 1
                End If
                End If
                
                If C_enum_contains_space = True And C_Enumname <> "Not Applicable" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "Enumeration Name must have underscore for each word it contains."
                    Sheets("Attribute").Cells(17, 3).Interior.Color = RGB(255, 0, 0)
                    lastARow = lastARow + 1
                    Else
                    
                End If
                
                If C_Descriptive = True And C_Identifier = True Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Descriptive and Opaque Identifier are disjoint and cannot both be true."
                Sheets("Attribute").Cells(30, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
                
           End If
           
           If C_subtype = True And C_datatype <> "SymbolVariant@1.0" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "A subtype attribute must be modeled with a SymbolVariant datatype."
                Sheets("Attribute").Cells(37, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
                
           End If
           
            If C_subtype = True And C_hassubtypecls < 1 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "A subtype attribute must have an attribute class of subtype_attribute."
                lastARow = lastARow + 1
                Else
                
           End If
                                
           
           If C_Branded = True And C_Identifier = True Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Branded and Opaque Identifier are disjoint and cannot both be true."
                Sheets("Attribute").Cells(31, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
               
           End If
           
           If C_Branded = True And C_Descriptive = True Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Branded and Descriptive are disjoint and cannot both be true."
                Sheets("Attribute").Cells(29, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
            Else
               
           End If
           
            If C_Internal = "True" And c_listing_group = "Not Applicable" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "If attribute is on listing template then a listing group is required."
                    Sheets("Attribute").Cells(28, 4).Interior.Color = RGB(255, 0, 0)
                    lastARow = lastARow + 1
            End If
            
             
            If C_Internal = "False" And c_listing_group <> "Not Applicable" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "If attribute is not on listing template then listing group should be Not Applicable."
                    Sheets("Attribute").Cells(28, 4).Interior.Color = RGB(255, 0, 0)
                    lastARow = lastARow + 1
            End If
            
            If Sheets("Attribute Properties").Cells(2, 5) = "value" And Sheets("Attribute Properties").Cells(2, 6) <> "visible_editable" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "Attribute with simple datatype requires value variant to be visible and editable."
                    lastARow = lastARow + 1
            End If
            
                
                
                If C_Identifier = "Not Applicable" Then
                    Sheets("attribute").Cells(32, 3).Value = "FALSE"
                End If
           
                If C_Identifier = "Not Applicable" Then
                    Sheets("attribute").Cells(30, 3).Value = "FALSE"
                End If
            
                If C_Descriptive = "Not Applicable" Then
                    Sheets("attribute").Cells(31, 3).Value = "FALSE"
                End If
           
                If C_Identity = "Not Applicable" Then
                    Sheets("attribute").Cells(28, 3).Value = "FALSE"
                End If
                
                If C_changetype <> "Component Attribute" Then
                    Sheets("attribute").Cells(29, 3).Value = "FALSE"
                End If
                           
                If C_Internal = "Not Applicable" Then
                    Sheets("attribute").Cells(27, 3).Value = "FALSE"
                End If
                
                If C_Branded = "Not Applicable" Then
                    Sheets("attribute").Cells(33, 3).Value = "FALSE"
                End If
                              
                If C_configurable = "Not Applicable" Then
                    Sheets("attribute").Cells(35, 3).Value = "FALSE"
                End If
           
                If C_refinement = "Not Applicable" Then
                    Sheets("attribute").Cells(34, 3).Value = "FALSE"
                End If
                
                                             
                               
           If C_changetype = "Component Attribute" Or C_changetype = "Complex Attribute With Sub-attribute" Then
        
            If C_contains_unit = 0 Then
                If C_subdatatype = "UnitDecimalStruct@1.0" Or C_datatype = "NormalizedUnitDecimalStruct@1.0" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "An enumeration constraint must exist for the Unit for a unit Decimal Datatype"
                lastARow = lastARow + 1
                End If
                End If
           
           If C_subdatatype = "SymbolStruct@1.0" And C_contains_enum = 0 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "An enumeration constraint must exist for each subattribute that has a SymbolStruct datatype."
                lastARow = lastARow + 1
           End If
           
           If C_subdescriptive = True And C_subidentifier = True Then
                Sheets("Errors").Cells(lastARow, 1).Value = "Descriptive and Opaque Identifier are disjoint and cannot both be true"
                Sheets("subattribute").Cells(20, 2).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
                Else
                
           End If
           
           If C_subbranded = True And C_subidentifier = True Then
               Sheets("Errors").Cells(lastARow, 1).Value = "Branded and Opaque Identifier are disjoint and cannot both be true on sub-attribute."
               Sheets("subattribute").Cells(21, 2).Interior.Color = RGB(255, 0, 0)
               lastARow = lastARow + 1
               Else
               
           End If
           
           If C_subbranded = True And C_subdescriptive = True Then
               Sheets("Errors").Cells(lastARow, 1).Value = "Branded and Descriptive are disjoint and cannot both be true on sub-attribute."
               Sheets("subattribute").Cells(20, 2).Interior.Color = RGB(255, 0, 0)
               lastARow = lastARow + 1
               Else
              
           End If
              
           If C_subreplacement = True And C_subreplaced = "" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If this is a replacement attribute, fill attribute to be replaced in cell D34 on sub-attribute. "
                Sheets("subattribute").Cells(25, 3).Interior.Color = RGB(255, 0, 0)
                lastARow = lastARow + 1
               Else
              
           End If
                    
           If C_contains_scale = 0 Then
                If C_subdatatype = "UnitDecimalStruct@1.0" Or C_datatype = "NormalizedUnitDecimalStruct@1.0" Or C_datatype = "DecimalStuct@1.0" Or C_datatype = "Price@1.0" Or C_datatype = "FractionMeasure@1.0" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "A max_scale constraint must exist for a decimal based datatype"
                lastARow = lastARow + 1
                End If
           End If
           
           If C_contains_maxprecision = 0 Then
                If C_subdatatype = "UnitDecimalStruct@1.0" Or C_datatype = "NormalizedUnitDecimalStruct@1.0" Or C_datatype = "DecimalStuct@1.0" Or C_datatype = "Price@1.0" Or C_datatype = "FractionMeasure@1.0" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "A max_precision constraint must exist for a decimal based datatype"
                lastARow = lastARow + 1
                End If
           End If
           
           If C_contains_minval = 0 Then
                If C_subdatatype = "UnitDecimalStruct@1.0" Or C_datatype = "NormalizedUnitDecimalStruct@1.0" Or C_datatype = "DecimalStuct@1.0" Or C_datatype = "Price@1.0" Or C_datatype = "FractionMeasure@1.0" Or C_datatype = "IntStruct@1.0" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "A min_value constraint must exist for a numeric datatype"
                lastARow = lastARow + 1
                End If
           End If
           
           If C_contains_maxval = 0 Then
                If C_subdatatype = "UnitDecimalStruct@1.0" Or C_datatype = "NormalizedUnitDecimalStruct@1.0" Or C_datatype = "DecimalStuct@1.0" Or C_datatype = "Price@1.0" Or C_datatype = "FractionMeasure@1.0" Or C_datatype = "IntStruct@1.0" Then
                Sheets("Errors").Cells(lastARow, 1).Value = "A max_value constraint must exist for a numeric datatype"
                lastARow = lastARow + 1
                End If
           End If
           
            If C_subdatatype = "SymbolStruct@1.0" And C_subenum_name = "Not Applicable" Then
                    Sheets("subattribute").Cells(7, 2).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "Not Applicable is not a valid enumeration name for a symbol struct on sub-attribute."
                    lastARow = lastARow + 1
                Else
                   
            End If
            
               If C_subrefinement = True Then
                    If C_contains_eligible = 0 Or C_contains_eligible = 0 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If Is Refinement Candidate = True then refinement eligible or refinement partially eligible attribute class must be added on Cust Rel Requirements Tab."
                lastARow = lastARow + 1
                Else
                End If
                End If
                
                If C_subrefinement = False Then
                    If C_contains_ineligible = 0 Then
                Sheets("Errors").Cells(lastARow, 1).Value = "If Is Refinement Candidate = False then refinement ineligible attribute class must be added on Cust Rel Requirements Tab."
                lastARow = lastARow + 1
                Else
                End If
                End If
             
                                 
                
            If C_subenum_name <> "Not Applicable" And C_subisenum = "False" Then
                    Sheets("subattribute").Cells(7, 2).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "If attribute is not enumerated, enumeration name should equal Not Applicable on sub-attribute."
                    lastARow = lastARow + 1
                Else
                    
            End If
            
             If C_subenum_space = "true" Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "Enumeration Name must have underscore for each word it contains."
                    Sheets("subattribute").Cells(7, 2).Interior.Color = RGB(255, 0, 0)
                    lastARow = lastARow + 1
                    Else
                    
             End If
             
             
            If C_subconcept = True Then
                    Sheets("Errors").Cells(lastARow, 1).Value = "Concept name should not contain an underscore on sub-attribute."
                    Sheets("subattribute").Cells(3, 2).Interior.Color = RGB(255, 0, 0)
                    lastARow = lastARow + 1
                    Else
                    
             End If
           
             
             If C_subdatatype = "SymbolStruct@1.0" And C_subenumtype <> "Closed" Then
                    Sheets("subattribute").Cells(7, 3).Interior.Color = RGB(255, 0, 0)
                    Sheets("Errors").Cells(lastARow, 1).Value = "Enumeration for symbol struct attribute must be closed on sub-attribute."
                    lastARow = lastARow + 1
             Else
                   
             End If
                     
          
      End If
                      
       
    Next iRow
    
End With
   
Application.Calculation = xlAutomatic
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
End Sub
Sub Refresh_Click()
Dim Answer As VbMsgBoxResult

    Answer = MsgBox("Do you want to refresh constraints? This will erase all current entries", vbYesNo)
    If Answer = vbYes Then

C_datatype = Worksheets("Attribute").Cells(15, 3).Value
C_changetype = Worksheets("Attribute").Cells(4, 3).Value
C_subdatatype = Worksheets("subattribute").Cells(5, 2).Value
C_enumeration = Worksheets("Attribute").Cells(17, 3).Value
C_subenum_name = Worksheets("subattribute").Cells(7, 2).Value

Application.Calculation = xlManual
Application.DisplayStatusBar = False
Application.ScreenUpdating = False

  Sheets("constraint").Range("A2:I20").ClearContents

  If C_datatype = "Complex Data Type" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
            
  End If
  
If C_changetype = "New Simple Attribute" Then

 If C_datatype = "SymbolVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 8).Value = C_enumeration
    Sheets("constraint").Cells(3, 9).Value = "Insert"
   
 End If
 
 If C_datatype = "BoolVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 8).Value = "boolean"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
   
  End If
 
  If C_datatype = "PriceVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "max_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "min_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "currency"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "enumeration"
    Sheets("constraint").Cells(7, 8).Value = "currency_code"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
 
 End If
 
   If C_datatype = "TypedPriceVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "min_value"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_scale"
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "value"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "max_precision"
    Sheets("constraint").Cells(7, 8).Value = 12
    Sheets("constraint").Cells(7, 9).Value = "Insert"
    Sheets("constraint").Cells(8, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(8, 2).Value = "currency"
    Sheets("constraint").Cells(8, 3).Value = "10480"
    Sheets("constraint").Cells(8, 4).Value = "contributor"
    Sheets("constraint").Cells(8, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(8, 6).Value = "enumeration"
    Sheets("constraint").Cells(8, 8).Value = "currency_code"
    Sheets("constraint").Cells(8, 9).Value = "Insert"
 
 End If
  
  
If C_datatype = "LocalizedStringVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "max_length"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    If C_enumeration <> "Not Applicable" Then
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "soft_enumeration"
    Sheets("constraint").Cells(4, 8).Value = C_enumeration
    Sheets("constraint").Cells(3, 8).Value = "50"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    End If
End If

If C_datatype = "StringVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "max_length"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
End If
 
If C_datatype = "UnitDecimalVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "unit"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "enumeration"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
    
End If

If C_datatype = "TypedUnitDecimalVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "min_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_value"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_scale"
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "value"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "max_precision"
    Sheets("constraint").Cells(7, 8).Value = 12
    Sheets("constraint").Cells(7, 9).Value = "Insert"
    Sheets("constraint").Cells(8, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(8, 2).Value = "unit"
    Sheets("constraint").Cells(8, 3).Value = "10480"
    Sheets("constraint").Cells(8, 4).Value = "contributor"
    Sheets("constraint").Cells(8, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(8, 6).Value = "enumeration"
    Sheets("constraint").Cells(8, 9).Value = "Insert"
    
End If

If C_datatype = "TypedIntRangeVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "minimum"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "min_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "minimum"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_value"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "maximum"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "min_value"
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "maximum"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "max_value"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
End If

If C_datatype = "UnitDecimalRangeVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "minimum.value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "minimum.value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "minimum.value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "minimum.value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "minimum.unit"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "enumeration"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
    Sheets("constraint").Cells(8, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(8, 2).Value = "maximum.value"
    Sheets("constraint").Cells(8, 3).Value = "10480"
    Sheets("constraint").Cells(8, 4).Value = "contributor"
    Sheets("constraint").Cells(8, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(8, 6).Value = "min_value"
    Sheets("constraint").Cells(8, 9).Value = "Insert"
    Sheets("constraint").Cells(9, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(9, 2).Value = "maximum.value"
    Sheets("constraint").Cells(9, 3).Value = "10480"
    Sheets("constraint").Cells(9, 4).Value = "contributor"
    Sheets("constraint").Cells(9, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(9, 6).Value = "max_value"
    Sheets("constraint").Cells(9, 9).Value = "Insert"
    Sheets("constraint").Cells(10, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(10, 2).Value = "maximum.value"
    Sheets("constraint").Cells(10, 3).Value = "10480"
    Sheets("constraint").Cells(10, 4).Value = "contributor"
    Sheets("constraint").Cells(10, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(10, 6).Value = "max_scale"
    Sheets("constraint").Cells(10, 9).Value = "Insert"
    Sheets("constraint").Cells(11, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(11, 2).Value = "maximum.value"
    Sheets("constraint").Cells(11, 3).Value = "10480"
    Sheets("constraint").Cells(11, 4).Value = "contributor"
    Sheets("constraint").Cells(11, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(11, 6).Value = "max_precision"
    Sheets("constraint").Cells(11, 8).Value = 12
    Sheets("constraint").Cells(11, 9).Value = "Insert"
    Sheets("constraint").Cells(12, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(12, 2).Value = "maximum.unit"
    Sheets("constraint").Cells(12, 3).Value = "10480"
    Sheets("constraint").Cells(12, 4).Value = "contributor"
    Sheets("constraint").Cells(12, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(12, 6).Value = "enumeration"
    Sheets("constraint").Cells(12, 9).Value = "Insert"
End If


If C_datatype = "UnitIntVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "unit"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "enumeration"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
End If

If C_datatype = "ItemLengthWidthHeight@1.0" Then
    
    Call itemlengthwidthheight

End If

If C_datatype = "ItemLengthWidth@1.0" Then
    
    Call ItemLengthWidth

End If

If C_datatype = "ItemWidthHeight@1.0" Then
    
    Call itemwidthheight

End If


If C_datatype = "ItemDiameterLength@1.0" Then
    
    Call itemdiameterlength

End If

If C_datatype = "ItemDiameterThickness@1.0" Then
    
    Call itemdiameterthickness

End If


If C_datatype = "ItemDepthWidthHeight@1.0" Then
    
    Call itemdepthwidthheight

End If

If C_datatype = "ItemWidthHeightThickness@1.0" Then
    
    Call itemwidthheightthickness

End If


If C_datatype = "ItemWidthDiameterHeight@1.0" Then
    
    Call itemwidthdiameterheight

End If

If C_datatype = "ItemLengthWidthThickness@1.0" Then
    
    Call itemlengthwidththickness

End If

If C_datatype = "ItemHeightThickness@1.0" Then
    
    Call itemheightthickness

End If

If C_datatype = "ItemLengthWidthDepth@1.0" Then
    
    Call itemlengthwidthdepth
    
End If

If C_datatype = "DecimalVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    
End If
 
If C_datatype = "IntVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
   
End If

If C_datatype = "FractionMeasureVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "decimal_value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "decimal_value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "decimal_value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "decimal_value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "unit"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "enumeration"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
    Sheets("constraint").Cells(8, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(8, 2).Value = "string_value"
    Sheets("constraint").Cells(8, 3).Value = "10480"
    Sheets("constraint").Cells(8, 4).Value = "contributor"
    Sheets("constraint").Cells(8, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(8, 6).Value = "max_length"
    Sheets("constraint").Cells(8, 9).Value = "Insert"
End If

If C_datatype = "FractionVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "decimal_value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "decimal_value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "decimal_value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "decimal_value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "string_value"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "max_length"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
End If

If C_datatype = "TypedStringVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_length"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
End If

If C_datatype = "TypedSymbolVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "enumeration"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
End If

If C_datatype = "TypedIntVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "min_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_value"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
End If

If C_datatype = "StringListVariant@1.0" Or C_datatype = "SymbolListVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
End If

If C_datatype = "TypeSubTypeVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "sub_type"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "enumeration"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
End If

If C_datatype = "NormalizedUnitDecimalVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "value"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "min_value"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "value"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "max_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "value"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_scale"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
    Sheets("constraint").Cells(6, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(6, 2).Value = "value"
    Sheets("constraint").Cells(6, 3).Value = "10480"
    Sheets("constraint").Cells(6, 4).Value = "contributor"
    Sheets("constraint").Cells(6, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(6, 6).Value = "max_precision"
    Sheets("constraint").Cells(6, 8).Value = 12
    Sheets("constraint").Cells(6, 9).Value = "Insert"
    Sheets("constraint").Cells(7, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(7, 2).Value = "unit"
    Sheets("constraint").Cells(7, 3).Value = "10480"
    Sheets("constraint").Cells(7, 4).Value = "contributor"
    Sheets("constraint").Cells(7, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(7, 6).Value = "enumeration"
    Sheets("constraint").Cells(7, 9).Value = "Insert"
    Sheets("constraint").Cells(8, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(8, 2).Value = "normalized_value.unit"
    Sheets("constraint").Cells(8, 3).Value = "10480"
    Sheets("constraint").Cells(8, 4).Value = "contributor"
    Sheets("constraint").Cells(8, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(8, 6).Value = "enumeration"
    Sheets("constraint").Cells(8, 9).Value = "Insert"
    Sheets("constraint").Cells(9, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(9, 2).Value = "normalized_value.value"
    Sheets("constraint").Cells(9, 3).Value = "10480"
    Sheets("constraint").Cells(9, 4).Value = "contributor"
    Sheets("constraint").Cells(9, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(9, 6).Value = "min_value"
    Sheets("constraint").Cells(9, 9).Value = "Insert"
    Sheets("constraint").Cells(10, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(10, 2).Value = "normalized_value.value"
    Sheets("constraint").Cells(10, 3).Value = "10480"
    Sheets("constraint").Cells(10, 4).Value = "contributor"
    Sheets("constraint").Cells(10, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(10, 6).Value = "max_value"
    Sheets("constraint").Cells(10, 9).Value = "Insert"
    
End If





If C_datatype = "TimeStampVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(12, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
End If


If C_datatype = "TypedQuantityVariant@1.0" Then
    Sheets("constraint").Cells(2, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(2, 2).Value = "*"
    Sheets("constraint").Cells(2, 3).Value = "10480"
    Sheets("constraint").Cells(2, 4).Value = "contributor"
    Sheets("constraint").Cells(2, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(2, 6).Value = "max_occurs"
    Sheets("constraint").Cells(2, 8).Value = 1
    Sheets("constraint").Cells(2, 9).Value = "Insert"
    Sheets("constraint").Cells(3, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(3, 2).Value = "type"
    Sheets("constraint").Cells(3, 3).Value = "10480"
    Sheets("constraint").Cells(3, 4).Value = "contributor"
    Sheets("constraint").Cells(3, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(3, 6).Value = "enumeration"
    Sheets("constraint").Cells(3, 9).Value = "Insert"
    Sheets("constraint").Cells(4, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(4, 2).Value = "quantity"
    Sheets("constraint").Cells(4, 3).Value = "10480"
    Sheets("constraint").Cells(4, 4).Value = "contributor"
    Sheets("constraint").Cells(4, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(4, 6).Value = "min_value"
    Sheets("constraint").Cells(4, 9).Value = "Insert"
    Sheets("constraint").Cells(5, 1).Value = Worksheets("attribute").Cells(13, 3).Value
    Sheets("constraint").Cells(5, 2).Value = "quantity"
    Sheets("constraint").Cells(5, 3).Value = "10480"
    Sheets("constraint").Cells(5, 4).Value = "contributor"
    Sheets("constraint").Cells(5, 5).Value = "PRODUCT"
    Sheets("constraint").Cells(5, 6).Value = "max_value"
    Sheets("constraint").Cells(5, 9).Value = "Insert"
End If

End If

If C_changetype = "Component Attribute" Then

Call component
       
End If
End If

Application.Calculation = xlAutomatic
Application.DisplayStatusBar = True
Application.ScreenUpdating = True

End Sub
