Imports NBC.IR

Public Module VBUtils
    Public Function CreateMainModule(block As Statement) As Statement
        Dim st_sub = New Statement With {.Type = SType.FUNCTION_DECLARATION, .Identifier = "Main", .Statements = block.Statements}
        Dim st_mod = New Statement With {.Type = SType.MODULE_, .Identifier = "Main", .Statements = New List(Of Statement)}
        st_mod.Statements.Add(st_sub)
        Return st_mod
    End Function

    Public Sub CorrectForAssignment(InvokeExpression As Statement)
        If InvokeExpression.Type = SType.INVOKE_EXPRESSION Then
            Dim exp = InvokeExpression.ActualParameters(0)
            If exp.Type = ExpressionType.BINARY_OP_EQUALS Then
                Dim lhs = exp.Args(0)
                Dim rhs = exp.Args(1)
                InvokeExpression.Type = SType.ASSIGNMENT
                InvokeExpression.ActualParameters(0) = lhs
                InvokeExpression.ActualParameters.Add(rhs)
            End If
        End If
    End Sub

    Public Sub RewriteCompositeAssignment(InvokeExpression As Statement, AssignmentType As ExpressionType, rhs As Expression)
        If InvokeExpression.Type = SType.INVOKE_EXPRESSION Then
            InvokeExpression.Type = SType.COMPOSITE_ASSIGNMENT
            InvokeExpression.AssignmentType = AssignmentType
            InvokeExpression.ActualParameters.Add(rhs)
        End If
    End Sub

    Public Const NameOfPrint = "_system_print"
End Module
