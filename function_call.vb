Option Explicit
Declare Function stringBack Lib "libforexcel.dll" (ByVal thatString As String) As String

Public Function myFunction(ByVal test As String) As String
	Dim st As String
	st = stringBack(test)
	myFunction = st
End Function
