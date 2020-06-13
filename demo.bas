Attribute VB_Name = "RibbonX_Code"
'Entry point for RibbonX button click
Sub ShowATPDialog(control As IRibbonControl)
    Application.Run ("fDialog")
End Sub

'Callback for RibbonX button label
Sub GetATPLabel(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets("RES").Range("A10").Value
End Sub

