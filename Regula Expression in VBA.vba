Sub windowsDataWP()

Dim obj As Object
Dim RegExp As Object
Set RegExp = CreateObject("VBScript.RegExp")

Dim strPattern As String

RegExp.Pattern = "Effective\s+[0-9/]+\s+Distance.\s+([0-9.]+)\s+Effective"
RegExp.Global = True
RegExp.IgnoreCase = True

Dim dataobj As New MSForms.DataObject
Dim str As String
Set obj = CreateObject("wscript.shell")

Application.Wait (Now + TimeValue("00:00:01"))
rowCounter = 4
windowName = Sheets("Sheet1").Range("B1").Value
obj.AppActivate Trim(windowName)

Application.Wait (Now + TimeValue("00:00:02"))
zip1 = 1
Do While zip1 <> ""
    obj.AppActivate windowName
    zip1 = Sheets("Sheet1").Range("B" & rowCounter).Value
    zip1 = CStr(zip1)
    If Len(zip1) < 5 Then
        zip1 = "0" + zip1
    End If
    
    zip2 = Sheets("Sheet1").Range("C" & rowCounter).Value
    zip2 = CStr(zip2)
    If Len(zip2) < 5 Then
        zip2 = "0" + zip2
    End If
    
    obj.SendKeys "{TAB 4}"
    obj.SendKeys zip1
    obj.SendKeys "{TAB 2}"
    obj.SendKeys zip2
    obj.SendKeys "{ENTER}"
    
    Application.Wait (Now + TimeValue("00:00:02"))
    obj.SendKeys "%es"
    Application.Wait (Now + TimeValue("00:00:01"))
    obj.SendKeys "^c"
    Application.Wait (Now + TimeValue("00:00:01"))
    
    dataobj.GetFromClipboard
    
    str = dataobj.GetText
    
    Set allMatches = RegExp.Execute(str)
    result = ""
    
    If allMatches.count <> 0 Then
    
        result = allMatches.Item(0).SubMatches.Item(0)
    
    End If

    ExtractSDI = result
    
    Worksheets("Sheet1").Range("D" & rowCounter).Value = ExtractSDI
    
    rowCounter = rowCounter + 1
    
    zip1 = Sheets("Sheet1").Range("B" & rowCounter).Value
    
    Loop

End Sub
