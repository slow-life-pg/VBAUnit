Attribute VB_Name = "Bridge"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Const

Public Const ResultSheet As String = "MacroReturn"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public Procedures

'' CallByNameを使わないと、ByRefの値が戻ってこない。
'' Runは外からも呼べるが、Bridgeを経由する必要がある。
Public Function CallMacro(ByVal callobj As Variant, ByVal workbookname As String, ByVal macroname As String, ByRef params() As Variant) As Variant
    Dim callparams() As Variant
    
    ReDim callparams(UBound(params) + 3)
    
    Dim i As Long
    For i = 0 To UBound(params)
        If IsObject(params(i)) Then
            Set callparams(i + 1) = params(i)
        Else
            callparams(i + 1) = params(i)
        End If
    Next
    
    Err.Clear
    
    On Error GoTo MacroError
    
    callparams(UBound(params) + 2) = 0
    callparams(UBound(params) + 3) = ""
    
    If IsNull(callobj) Then
        Call CallApplicationMacro(callparams, GetMacroPath(workbookname, macroname))
    Else
        Call CallObjectMacro(callobj, callparams, GetMacroPath(workbookname, macroname))
    End If
    
    GoTo ExitFunction
    
MacroError:
    callparams(UBound(params) + 2) = Err.Number
    callparams(UBound(params) + 3) = Err.Description
    
ExitFunction:
    CallMacro = callparams
End Function

Public Sub Free(ByRef params As Variant)
    Erase params
End Sub

Public Function GetRegexp() As RegExp
    Set GetRegexp = New RegExp
End Function

Public Function GetNewCollection() As Variant
    Set GetNewCollection = New Collection
End Function

Public Function GetNewDictionary() As Variant
    Set GetNewDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Function GetShapeAction(ByVal bookName As String, ByVal sheetName As String, ByVal shapeName) As String
    Dim result As String
    
    Dim book As Workbook
    Dim i As Long
    
    Set book = Nothing
    For i = 1 To Application.Workbooks.Count
        If Application.Workbooks(i).Name = bookName Then
            Set book = Application.Workbooks(i)
            Exit For
        End If
    Next
    
    If book Is Nothing Then
        result = ""
        GoTo ExitFunction
    End If
    
    result = book.Worksheets(sheetName).Shapes(shapeName).OnAction
    
ExitFunction:
    GetShapeAction = result
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' [Region] Private Procedures

Private Function GetMacroPath(ByVal workbookname As String, ByVal macroname As String) As String
    GetMacroPath = "'" & workbookname & "'!" & macroname
End Function

Private Sub CallApplicationMacro(ByRef callparams() As Variant, ByVal macropath As String)
    Select Case UBound(callparams) - 2
        Case 0
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath)
        Case 1
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1))
        Case 2
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2))
        Case 3
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3))
        Case 4
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4))
        Case 5
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5))
        Case 6
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6))
        Case 7
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7))
        Case 8
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8))
        Case 9
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9))
        Case 10
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10))
        Case 11
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11))
        Case 12
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12))
        Case 13
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13))
        Case 14
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13), callparams(14))
        Case 15
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13), callparams(14), callparams(15))
        Case 16
            callparams(0) = CallByName(Application, "Run", VbMethod, macropath, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13), callparams(14), callparams(15), callparams(16))
    End Select
End Sub

Private Sub CallObjectMacro(ByVal callobj As Object, ByRef callparams() As Variant, ByVal macropath As String, Optional ByVal calltype As Long = VbMethod)
    Select Case UBound(callparams) - 3
        Case 0
            callparams(0) = CallByName(callobj, macropath, calltype)
        Case 1
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1))
        Case 2
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2))
        Case 3
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3))
        Case 4
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4))
        Case 5
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5))
        Case 6
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6))
        Case 7
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7))
        Case 8
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8))
        Case 9
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9))
        Case 10
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10))
        Case 11
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11))
        Case 12
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12))
        Case 13
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13))
        Case 14
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13), callparams(14))
        Case 15
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13), callparams(14), callparams(15))
        Case 16
            callparams(0) = CallByName(callobj, macropath, calltype, callparams(1), callparams(2), callparams(3), callparams(4), callparams(5), callparams(6), callparams(7), callparams(8), callparams(9), callparams(10), callparams(11), callparams(12), callparams(13), callparams(14), callparams(15), callparams(16))
    End Select
End Sub

