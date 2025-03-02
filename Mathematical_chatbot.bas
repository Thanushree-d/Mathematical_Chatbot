Attribute VB_Name = "Module1"
Sub RunMathChatbot()
Attribute RunMathChatbot.VB_Description = "enable content"
Attribute RunMathChatbot.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim InputCell As Range
    Dim OutputCell As Range
    Dim Query As String
    Dim Response As String

    ' Define input and output cells
    Set InputCell = ThisWorkbook.Sheets(1).Range("A1") ' User query cell
    Set OutputCell = ThisWorkbook.Sheets(1).Range("B1") ' Chatbot response cell

    ' Get the user's query
    Query = InputCell.Value

    ' Call the chatbot function
    Response = MathChatbot(Query)

    ' Output the response
    OutputCell.Value = Response
End Sub

Function MathChatbot(Query As String) As String
    Dim Expression As String
    Dim Result As Double
    Dim Operation As String
    Dim Num1 As Double, Num2 As Double
    
    On Error GoTo ErrorHandler

    ' Process the query
    Query = LCase(Query) ' Convert to lowercase for uniformity
    
    ' Extract numbers and operations
    If InStr(Query, "plus") > 0 Then
        Operation = "+"
    ElseIf InStr(Query, "add") > 0 Then
        Operation = "+"
    ElseIf InStr(Query, "sum") > 0 Then
        Operation = "+"
    ElseIf InStr(Query, "minus") > 0 Then
        Operation = "-"
    ElseIf InStr(Query, "subtract") > 0 Then
        Operation = "-"
    ElseIf InStr(Query, "multiply") > 0 Then
        Operation = "*"
    ElseIf InStr(Query, "times") > 0 Then
        Operation = "*"
    ElseIf InStr(Query, "divide") > 0 Then
        Operation = "/"
    Else
        MathChatbot = "I didn't understand your question. Please use basic math operations like add, subtract, multiply, or divide."
        Exit Function
    End If

    ' Extract the numbers
    Dim Matches As Object
    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    Regex.Pattern = "\d+"
    Regex.Global = True
    
    Set Matches = Regex.Execute(Query)
    If Matches.Count >= 2 Then
        Num1 = CDbl(Matches(0))
        Num2 = CDbl(Matches(1))
    Else
        MathChatbot = "Please provide two numbers for the operation."
        Exit Function
    End If

    ' Formulate the expression and calculate
    Expression = Num1 & " " & Operation & " " & Num2
    Result = Application.Evaluate(Expression)
    
    ' Return the result
    MathChatbot = "The result is " & Result & "."
    Exit Function

ErrorHandler:
    MathChatbot = "There was an error processing your query. Please check your input."
End Function
