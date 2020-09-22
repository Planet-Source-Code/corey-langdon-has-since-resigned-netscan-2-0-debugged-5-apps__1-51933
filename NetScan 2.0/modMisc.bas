Attribute VB_Name = "modMisc"
Function Ping(Optional guid As String) As String
    'Utils: Sends a Ping Message
    '
    '-------
    If guid = "" Then guid = cx(guid)
    '-------
    Ping = guid & String(7, Chr(0))
End Function

Function cx$(h$)
'return actual characters from hex string
Dim a As Integer
Dim hs1$
Dim hs2$
Dim c$
Dim v As Integer
Dim o$
Dim msb As Integer, lsb As Integer
For a = 1 To Len(h$) Step 2
hs1$ = Mid$(h$, a, 1)
hs2$ = Mid$(h$, a + 1, 1)
c$ = LCase$(hs1$)
v = Val(c$)
If c$ = "a" Then v = 10
If c$ = "b" Then v = 11
If c$ = "c" Then v = 12
If c$ = "d" Then v = 13
If c$ = "e" Then v = 14
If c$ = "f" Then v = 15
msb = v
c$ = LCase$(hs2$)
v = Val(c$)
If c$ = "a" Then v = 10
If c$ = "b" Then v = 11
If c$ = "c" Then v = 12
If c$ = "d" Then v = 13
If c$ = "e" Then v = 14
If c$ = "f" Then v = 15
lsb = v
o$ = o$ & Chr$(msb * &H10 + lsb)
Next
cx$ = o$
End Function

