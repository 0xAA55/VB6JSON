Attribute VB_Name = "modTests"
Option Explicit

Sub Test1()
Dim JsonObj As Variant
ParseJSONString2 " [ 114, 514, ""1919"", 8.10e2 ] ", JsonObj

Debug.Print JsonObj(0); JsonObj(1); JsonObj(2); JsonObj(3)
End Sub

Sub Test2()
Dim JsonObj As Variant
ParseJSONString2 "{""Mike"": 123, ""Mary"": 456, ""Sam"": 789}", JsonObj

Debug.Print "Sam: "; JsonObj("Sam")
Debug.Print "Mike: "; JsonObj("Mike")
Debug.Print "Mary: "; JsonObj("Mary")
End Sub

Sub Test3()
Dim JsonObj As Variant
ParseJSONString2 "{""success"": true, ""data"": {""text"": ""Hello Json!"", ""title"": ""VB6 Json""}}", JsonObj
'ParseJSONString2 "{""success"": false, ""wording"": ""API failed.""}", JsonObj

If JsonObj("success") = False Then
    MsgBox JsonObj("wording"), vbExclamation, "API ∑µªÿ ß∞‹°£"
Else
    MsgBox JsonObj("data")("text"), vbInformation, JsonObj("data")("title")
End If
End Sub

Sub Test4()
Debug.Print JSONToString(ParseJSONString("""\ud800"""))
End Sub
