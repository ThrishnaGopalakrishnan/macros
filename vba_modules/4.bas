Attribute VB_Name = "Module4"
Sub datatype_Click()
Dim Answer As VbMsgBoxResult

    Answer = MsgBox("Do you want to refresh datatype sdl? This will erase all current entries", vbYesNo)
    If Answer = vbYes Then

C_datatype = Worksheets("Attribute").Cells(14, 3).Value
C_changetype = Worksheets("Attribute").Cells(3, 3).Value
C_subenum_name = Worksheets("subattribute").Cells(7, 2).Value

Application.Calculation = xlManual
Application.DisplayStatusBar = False
Application.ScreenUpdating = False

  Sheets("datatype").Cells(13, 1).ClearContents
  
 If C_changetype = "New Simple Attribute" Then
 Sheets("datatype").Cells(13, 1).Value = "Not Applicable"
 Else
  If C_changetype = "New Complex Attribute" Then
   Sheets("datatype").Cells(13, 1).Value = "{"&CHAR(10)&"ccp_datatype_name:["&CHAR(10)& Datatype!A3&","&CHAR(10)&"definition:type::{"&CHAR(10)&"base: '"&Datatype!$B$1&","&CHAR(10)&"name:"&Datatype!$A$3&","&CHAR(10)&"major_version:1,"&CHAR(10)&"minor_version:0,"&CHAR(10)&"fields:["&CHAR(10)&"{"&CHAR(10)&"name:"&Datatype!$C$3&","&CHAR(10)&"type: '"&Datatype!$D$3&","&CHAR(10)&"min_occurs:"&Datatype!$F$3&","&CHAR(10)&"max_occurs:"&Datatype!$G$3&","&CHAR(10)&"nullable:false"&CHAR(10)&"}"&","&CHAR(10)&"{"&CHAR(10)&"name:"&Datatype!$C$4&","&CHAR(10)&"type: '"&Datatype!$D$4&","&CHAR(10)&"min_occurs:"&Datatype!$F$4&","&CHAR(10)&"max_occurs:"&Datatype!$G$4&","&CHAR(10)&"nullable:false"&CHAR(10)&"}"&CHAR(10)&"{"&CHAR(10)&"name:"&Datatype!$C$5&","&CHAR(10)&"type: '"&Datatype!$D$5&","&CHAR(10)&"min_occurs:"&Datatype!$F$5&","&CHAR(10)&"max_occurs:"&Datatype!$G$5&","&CHAR(10)&"nullable:false"&CHAR(10)&"}"&CHAR(10)&"]"&CHAR(10)&"]"&CHAR(10)&"}""
End If
End If
Application.Calculation = xlAutomatic
Application.DisplayStatusBar = True
Application.ScreenUpdating = True

End Sub
