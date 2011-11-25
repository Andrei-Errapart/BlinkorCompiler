Imports System.Text

Namespace IR
    Public Module IRUtils
        Public Function ConcatStrings(Array As IEnumerable(Of String), ByVal Separator As String) As String
            Dim sb = New StringBuilder
            Dim is_first = True
            For Each s In Array
                If is_first Then
                    is_first = True
                Else
                    sb.Append(Separator)
                End If
                sb.Append(s)
            Next
            Return sb.ToString()
        End Function
        Public Function ConcatObjects(Array As IEnumerable(Of Object), ByVal Separator As String) As String
            Dim sb = New StringBuilder
            Dim is_first = True
            For Each o In Array
                If is_first Then
                    is_first = True
                Else
                    sb.Append(Separator)
                End If
                sb.Append(o.ToString())
            Next
            Return sb.ToString()
        End Function

        Public Function StringOfAccessModifier(am As AccessModifier, suffix As String) As String
            Select Case am
                Case AccessModifier.NONE
                    Return ""
                Case AccessModifier.PUBLIC_
                    Return "Public" & suffix
                Case AccessModifier.PROTECTED_
                    Return "Protected" & suffix
                Case AccessModifier.FRIEND_
                    Return "Friend" & suffix
                Case AccessModifier.PRIVATE_
                    Return "Private" & suffix
                Case Else
                    Throw New ApplicationException("StringOfAccessModifier: Unknown case:" & am)
            End Select
        End Function

        Public Function StringOfModifier(m As ProcedureModifier) As String
            Select Case m
                Case ProcedureModifier.NONE
                    Return ""
                Case ProcedureModifier.SHADOWS_
                    Return "Shadows"
                Case ProcedureModifier.SHARED_
                    Return "Shared"
                Case ProcedureModifier.OVERRIDABLE_
                    Return "Overridable"
                Case ProcedureModifier.NOTOVERRIDABLE_
                    Return "NotOverridable"
                Case ProcedureModifier.MUSTOVERRIDE_
                    Return "MustOverride"
                Case ProcedureModifier.OVERRIDES_
                    Return "Overrides"
                Case ProcedureModifier.OVERLOADS_
                    Return "Overloads"
                Case ProcedureModifier.DEFAULT_
                    Return "Default"
                Case ProcedureModifier.READONLY_
                    Return "ReadOnly"
                Case ProcedureModifier.WRITEONLY_
                    Return "WriteOnly"
                Case ProcedureModifier.CONSTRUCTOR
                    Return "Constructor"
                Case Else
                    Throw New ApplicationException("StringOfModifier: Unknown case:" & m)
            End Select
        End Function
        Public Function StringOfModifiers(Array As IEnumerable(Of ProcedureModifier)) As String
            Const Separator = " "
            Dim sb = New StringBuilder
            For Each m In Array
                sb.Append(StringOfModifier(m))
                sb.Append(Separator)
            Next
            Return sb.ToString()
        End Function
    End Module

    Public Enum ExpressionType
        ' Terminals
        STRING_  ' StringLiteral
        INTEGER_ ' IntegerLiteral
        VARIABLE_ACCESS     ' VariableName
        ' Binary operations
        BINARY_OP_PLUS      ' Args(0) + Args(1)
        BINARY_OP_MINUS     ' Args(0) - Args(1)
        BINARY_OP_MULTIPLY  ' Args(0) * Args(1)
        BINARY_OP_DIVIDE    ' Args(0) / Args(1)
        BINARY_OP_INT_DIVIDE ' Args(0) \ Args(1)
        BINARY_OP_POWER     ' Args(0) ^ Args(1)
        BINARY_OP_SHIFT_LEFT    ' Args(0) << Args(1)
        BINARY_OP_SHIFT_RIGHT   ' Args(0) >> Args(1)
        BINARY_OP_CONCAT    ' Args(0) & Args(1)
        BINARY_OP_MODULUS           ' Args(0) Mod Args(1)
        ' Comparisions.
        BINARY_OP_EQUALS    ' Args(0) = Args(1)
        BINARY_OP_NOT_EQUALS        ' Args(0) <> Args(1)
        BINARY_OP_LESSER            ' Args(0) < Args(1)
        BINARY_OP_LESSER_OR_EQUAL   ' Args(0) <= Args(1)
        BINARY_OP_GREATER           ' Args(0) > Args(1)
        BINARY_OP_GREATER_OR_EQUAL  ' Args(0) >= Args(1)
        'BINARY_OP_TYPEOF_IS         ' TypeOf Args(0) Is Args(1)
        BINARY_OP_IS            ' Args(0) Is Args(1)
        BINARY_OP_ISNOT         ' Args(0) IsNot Args(1)
        'BINARY_OP_LIKE      ' Args(0) Like Args(1)
        ' Boolean logic
        BINARY_OP_AND
        BINARY_OP_ANDALSO
        BINARY_OP_OR
        BINARY_OP_ORELSE
        BINARY_OP_XOR
        ' Unary operations
        UNARY_OP_MINUS      ' - Args(0)
        UNARY_OP_NOT        ' Not Args(0)
        ' 
        ' TODO: how to distinguish between array access and function call?
        ARRAY_ACCESS        ' ExpressionValue ( Args )
        FUNCTION_CALL       ' ExpressionValue ( Args )
        MEMBER_ACCESS       ' ExpressionValue "." StringValue
        NEW_                ' New NewType ( Args ) [ With { Initializers } ]
    End Enum

    Public Class Expression
        Public Type As ExpressionType
        Public Args As List(Of Expression)
        Public StringValue As String
        Public IntegerValue As Integer
        Public ExpressionValue As Expression
        Public NewType As VariableType
        Public Initializers As List(Of KeyValuePair(Of String, Expression))

        Public Overrides Function ToString() As String
            Dim sb As System.Text.StringBuilder
            Select Case Type
                Case ExpressionType.STRING_
                    Return """" & StringValue & """"
                Case ExpressionType.INTEGER_
                    Return IntegerValue.ToString()

                Case ExpressionType.BINARY_OP_PLUS
                    Return "(" & Args(0).ToString() & ") + (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_MINUS
                    Return "(" & Args(0).ToString() & ") - (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_MULTIPLY
                    Return "(" & Args(0).ToString() & ") * (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_DIVIDE
                    Return "(" & Args(0).ToString() & ") / (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_INT_DIVIDE
                    Return "(" & Args(0).ToString() & ") \ (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_POWER
                    Return "(" & Args(0).ToString() & ") ^ (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_SHIFT_LEFT
                    Return "(" & Args(0).ToString() & ") << (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_SHIFT_RIGHT
                    Return "(" & Args(0).ToString() & ") >> (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_CONCAT
                    Return "(" & Args(0).ToString() & ") & (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_MODULUS
                    Return "(" & Args(0).ToString() & ") Mod (" & Args(1).ToString() & ")"

                Case ExpressionType.BINARY_OP_EQUALS
                    Return "(" & Args(0).ToString() & ") = (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_NOT_EQUALS
                    Return "(" & Args(0).ToString() & ") <> (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_LESSER
                    Return "(" & Args(0).ToString() & ") < (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_LESSER_OR_EQUAL
                    Return "(" & Args(0).ToString() & ") <= (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_GREATER
                    Return "(" & Args(0).ToString() & ") > (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_GREATER_OR_EQUAL
                    Return "(" & Args(0).ToString() & ") >= (" & Args(1).ToString() & ")"

                Case ExpressionType.BINARY_OP_IS
                    Return "(" & Args(0).ToString() & ") Is (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_ISNOT
                    Return "(" & Args(0).ToString() & ") IsNot (" & Args(1).ToString() & ")"

                Case ExpressionType.BINARY_OP_AND
                    Return "(" & Args(0).ToString() & ") And (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_ANDALSO
                    Return "(" & Args(0).ToString() & ") AndAlso (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_OR
                    Return "(" & Args(0).ToString() & ") Or (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_ORELSE
                    Return "(" & Args(0).ToString() & ") OrElse (" & Args(1).ToString() & ")"
                Case ExpressionType.BINARY_OP_XOR
                    Return "(" & Args(0).ToString() & ") Xor (" & Args(1).ToString() & ")"

                Case ExpressionType.UNARY_OP_MINUS
                    Return " - " & Args(0).ToString()
                Case ExpressionType.VARIABLE_ACCESS
                    Return StringValue
                Case ExpressionType.ARRAY_ACCESS
                    Return "(" & Args(0).ToString() & ")" & "(" & Args(1).ToString() + ")"
                Case ExpressionType.FUNCTION_CALL
                    Dim s = "(" & ExpressionValue.ToString() & ")(" & IRUtils.ConcatObjects(Args, ", ") & ")"
                    Return s
                Case ExpressionType.MEMBER_ACCESS
                    Return "(" & ExpressionValue.ToString() & ")" & "." & StringValue
                Case ExpressionType.NEW_
                    sb = New StringBuilder()
                    sb.Append("New " & NewType.ToString())
                    If Args.Count > 0 Then
                        sb.Append("(" & IRUtils.ConcatObjects(Args, ", ") & ")")
                    End If
                    If Initializers IsNot Nothing AndAlso Initializers.Count > 0 Then
                        sb.Append(" With {")
                        For i = 0 To Initializers.Count - 1
                            Dim ii = Initializers(i)
                            If i > 0 Then
                                sb.Append(", ")
                            End If
                            sb.Append("." & ii.Key & " = " & ii.Value.ToString())
                        Next
                        sb.Append("}")
                    End If
                    Return sb.ToString()
                Case Else
                    Return "Expression.ToString(): Unknown type: " & (Type.ToString())
            End Select
        End Function
    End Class

    Public Structure Position
        Public PosInFile As Integer
        Public LineNo As Integer
        Public Column As Integer
    End Structure

    ' Name of the type, optionally with the rank of an array.
    ' This is purely a Basic fenomen.
    Public Class VBTypename
        Public Position As Position
        Public Name As String
        Public Ranks As List(Of List(Of Integer))   ' -1 = unlimited dimension.
        Public Overrides Function ToString() As String
            If Ranks Is Nothing Then
                Return Name
            Else
                Dim sb = New StringBuilder
                sb.Append(Name)
                For Each rr In Ranks
                    sb.Append("(")
                    For rr_index = 0 To rr.Count - 1
                        Dim rank = rr(rr_index)
                        If rank >= 0 Then
                            sb.Append(rank)
                        End If
                        If rr_index + 1 < rr.Count Then
                            sb.Append(", ")
                        End If
                    Next
                    sb.Append(")")
                Next
                Return sb.ToString()
            End If
        End Function
    End Class

    Public Enum VTType
        NAMED
        ARRAY
        STILL_UNKNOWN
    End Enum

    Public Enum AccessModifier
        NONE
        PUBLIC_
        PROTECTED_
        FRIEND_
        PRIVATE_
    End Enum
    Public Enum ProcedureModifier
        NONE
        SHADOWS_
        SHARED_
        OVERRIDABLE_
        NOTOVERRIDABLE_
        MUSTOVERRIDE_
        OVERRIDES_
        OVERLOADS_
        DEFAULT_
        READONLY_
        WRITEONLY_
        CONSTRUCTOR
    End Enum
    ' Type of variable, possibly nested.
    Public Class VariableType
        Public Type As VTType
        Public Name As String
        Public Ranks As List(Of Integer)
        ' Template parameters or nested type.
        Public NestedType As List(Of VariableType)

        Public Sub SetName(NewName As String)
            Name = NewName
            Type = VTType.NAMED
        End Sub

        Public Overrides Function ToString() As String
            Select Case Type
                Case VTType.NAMED
                    If NestedType IsNot Nothing AndAlso NestedType.Count > 0 Then
                        Return Name & "(Of " & IRUtils.ConcatObjects(NestedType, ",") & ")"
                    Else
                        Return Name
                    End If
                Case VTType.ARRAY
                    Dim sb = New StringBuilder
                    sb.Append("Array ([")
                    For RankIndex = 0 To Ranks.Count - 1
                        If RankIndex > 0 Then
                            sb.Append(", ")
                        End If
                        If Ranks(RankIndex) >= 0 Then
                            sb.Append(Ranks(RankIndex))
                        End If
                    Next
                    sb.Append("] Of " + NestedType(0).ToString() + ")")
                    Return sb.ToString()
                Case VTType.STILL_UNKNOWN
                    Return "Still Unknown"
            End Select
            Throw New ApplicationException("VariableType.ToString: Invalid type: " & Type.ToString())
        End Function
    End Class

    Public Enum VariableModifier
        NONE
        BYREF_
        BYVAL_
        CONST_
    End Enum

    ' Variable, Parameter or something else.
    Public Class Variable
        Public Position As Position
        Public Name As String
        Public Type As VariableType
        Public Init As Expression
        Public Modifier As VariableModifier

        ' Create from the VB typename.
        Public Shared Function CreateFromTypename(Typename As VBTypename) As Variable
            Dim r = New Variable With {.Name = Typename.Name, .Type = New VariableType}
            If (Typename.Ranks Is Nothing) OrElse (Typename.Ranks.Count = 0) Then
                r.Type.Type = VTType.STILL_UNKNOWN
            Else
                Dim CurrentType = r.Type
                For Each CurrentRanks In Typename.Ranks
                    CurrentType.Type = VTType.ARRAY
                    CurrentType.Ranks = CurrentRanks
                    CurrentType.NestedType = New List(Of VariableType)
                    CurrentType.NestedType.Add(New VariableType With {.Name = "", .Type = VTType.STILL_UNKNOWN})
                    CurrentType = CurrentType.NestedType(0)
                Next
                ' last one is still unknown
                CurrentType.Type = VTType.STILL_UNKNOWN
            End If
            Return r
        End Function

        ' Get the most nested type name.
        Public Function GetMostNestedType() As VariableType
            Dim r = Type
            While (r.NestedType IsNot Nothing) AndAlso (r.NestedType.Count > 0)
                r = r.NestedType(0)
            End While
            Return r
        End Function

        Public Sub SetMostNestedType(ByVal nt As VariableType)
            Dim r = Type
            Dim prev = Type
            While (r.NestedType IsNot Nothing) AndAlso (r.NestedType.Count > 0)
                prev = r
                r = r.NestedType(0)
            End While
            If prev Is r Then
                Me.Type = nt
            Else
                prev.NestedType(0) = nt
            End If
        End Sub

        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder
            Select Case Modifier
                Case VariableModifier.BYREF_
                    sb.Append("ByRef ")
                Case VariableModifier.BYVAL_
                    sb.Append("ByVal ")
            End Select
            sb.Append(Name)
            If Type.Type <> VTType.STILL_UNKNOWN Then
                sb.Append(" As " & Type.ToString())
            End If
            If Init IsNot Nothing Then
                sb.Append(" = " & Init.ToString())
            End If
            Return sb.ToString()
        End Function
    End Class

    Public Enum SType
        MULTIPLE_STATEMENTS     ' Statements
        VARIABLE_DECLARATION    ' Variables
        SUBROUTINE_CALL         ' Identifier(Args)
        INVOKE_EXPRESSION       ' Invoke ActualParameters(0)
        ASSIGNMENT              ' Args(0) = Args(1)
        COMPOSITE_ASSIGNMENT    ' Args(0) = Args(0) (AssignmentType) Args(1)
        IF_                     ' If Args(0) Then Statements(0) ElseIf Args(1) Statements(1) Else Statements(2) End If
        WHILE_                  ' While Args(0) Statements(0)
        SELECT_                 ' Select case Args(0) Case Args(1) statements(0) Case Args(2) Statements(1) .... Case Args(N-1) Statements(N-1) ... [Case Else Statements(N)]
        FOR_                    ' for Variables(0)=Args(0) to Args(1) ... Statements ... End For
        FOR_EACH                ' for Variables(0) in Args(0)  Statements
        EXIT_                   ' Exit Identifier (While/For/etc.)
        RETURN_                 ' Return ActualParameters(0)
        THROW_                  ' Throw ActualParameters(0)
        TRY_CATCH               ' Try Statements(0) (Catch Variable(0) Statements(1))*
        MODULE_                 ' Identifier, AccessModifier, Statements
        NAMESPACE_              ' Identifier, AccessModifier, Statements
        ENUM_                   ' Identifier, AccessModifier, Variables { .Name, .Init }
        CLASS_                  ' Identifier, AccessModifier, Statements
        FUNCTION_DECLARATION     ' Identifier, AccessModifier, FormalParameters, Statements, ReturnTypes
        OPTION_                 ' Args(0) as identifier, Args(1) as identifier
        IMPORTS_                ' FormalParameters.
    End Enum

    ' Statement is the building block.
    ' Every program is a list of (possibly nested) statements.
    Public Class Statement
        Public Type As SType
        Public Position As Position
        Public AccessModifier As AccessModifier = IR.AccessModifier.NONE
        Public Modifiers As List(Of ProcedureModifier)
        Public Statements As List(Of Statement)
        Public Identifier As String
        Public ActualParameters As List(Of Expression)  ' for Sub or Function call.
        Public FormalParameters As List(Of Variable)    ' For a Sub, Function or Try/Catch statement.
        Public Variables As List(Of Variable)           ' For 
        Public ReturnTypes As List(Of VariableType)     ' For FUNCTION_DECLARATION
        Public AssignmentType As ExpressionType         ' For COMPOSITE_ASSIGNMENT
        Public Const IndentationStep As String = "    "
        Public Sub ToStringBuilder(ByVal sb As System.Text.StringBuilder, Indent As String)
            ' indent
            Dim next_indent = Indent & IndentationStep

            Select Case Type
                Case SType.MULTIPLE_STATEMENTS
                    For Each st In Statements
                        st.ToStringBuilder(sb, Indent)
                    Next
                Case SType.VARIABLE_DECLARATION
                    For Each v In Variables
                        If v.Modifier = VariableModifier.CONST_ Then
                            sb.AppendLine(Indent & "Const " & v.ToString())
                        Else
                            sb.AppendLine(Indent & "Dim " & v.ToString())
                        End If
                    Next
                Case SType.ASSIGNMENT
                    sb.AppendLine(Indent & ActualParameters(0).ToString() & " = " & ActualParameters(1).ToString())
                Case SType.COMPOSITE_ASSIGNMENT
                    sb.Append(Indent & ActualParameters(0).ToString())
                    Select Case AssignmentType
                        Case ExpressionType.BINARY_OP_POWER
                            sb.Append(" &= ")
                        Case ExpressionType.BINARY_OP_MULTIPLY
                            sb.Append(" *= ")
                        Case ExpressionType.BINARY_OP_DIVIDE
                            sb.Append(" /= ")
                        Case ExpressionType.BINARY_OP_INT_DIVIDE
                            sb.Append(" \= ")
                        Case ExpressionType.BINARY_OP_PLUS
                            sb.Append(" += ")
                        Case ExpressionType.BINARY_OP_MINUS
                            sb.Append(" -= ")
                        Case ExpressionType.BINARY_OP_SHIFT_LEFT
                            sb.Append(" <<= ")
                        Case ExpressionType.BINARY_OP_SHIFT_RIGHT
                            sb.Append(" >>= ")
                        Case ExpressionType.BINARY_OP_CONCAT
                            sb.Append(" &= ")
                    End Select
                    sb.AppendLine(Indent & ActualParameters(1).ToString())
                Case SType.SUBROUTINE_CALL
                    sb.AppendLine(Indent & Identifier & "(" & IRUtils.ConcatObjects(ActualParameters, ",") & ")")
                Case SType.INVOKE_EXPRESSION
                    sb.AppendLine(Indent & ActualParameters(0).ToString())
                Case SType.IF_
                    sb.AppendLine(Indent & "If " & ActualParameters(0).ToString() & " Then")
                    Statements(0).ToStringBuilder(sb, next_indent)
                    For arg_index = 1 To ActualParameters.Count - 1
                        sb.AppendLine(Indent & "ElseIf " & ActualParameters(arg_index).ToString() & " Then")
                        Statements(arg_index).ToStringBuilder(sb, next_indent)
                    Next
                    If Statements.Count > ActualParameters.Count Then
                        sb.AppendLine(Indent & "Else")
                        Statements(ActualParameters.Count).ToStringBuilder(sb, next_indent)
                    End If
                    sb.AppendLine(Indent & "End If")
                Case SType.WHILE_
                    sb.AppendLine(Indent & "While " & ActualParameters(0).ToString())
                    Statements(0).ToStringBuilder(sb, next_indent)
                    sb.AppendLine(Indent & "End While")
                Case SType.SELECT_
                    sb.AppendLine(Indent & "Select Case " & ActualParameters(0).ToString())
                    Dim next2_indent = next_indent & IndentationStep
                    For i = 1 To ActualParameters.Count - 1
                        sb.AppendLine(next_indent & "Case " & ActualParameters(i).ToString())
                        Statements(i - 1).ToStringBuilder(sb, next2_indent)
                    Next
                    If Statements.Count = ActualParameters.Count Then
                        sb.AppendLine(next_indent & "Case Else")
                        Statements.Last().ToStringBuilder(sb, next2_indent)
                    End If
                    sb.AppendLine(Indent & "End Select")
                Case SType.FOR_
                    Dim v = Variables(0)
                    sb.AppendLine(Indent & "For " & v.ToString() & " To " & ActualParameters(0).ToString())
                    For Each st In Statements
                        st.ToStringBuilder(sb, next_indent)
                    Next
                    sb.AppendLine(Indent & "Next")
                Case SType.RETURN_
                    sb.AppendLine(Indent & "Return " & ActualParameters(0).ToString())
                Case SType.FOR_EACH
                    ' 1. First line.
                    Dim v = Variables(0)
                    sb.AppendLine(Indent & "For Each " & v.ToString() & " In " & ActualParameters(0).ToString())
                    ' 2. Middle lines.
                    For Each st In Statements
                        st.ToStringBuilder(sb, next_indent)
                    Next
                    ' 3. Last line.
                    sb.AppendLine(Indent & "Next")
                Case SType.EXIT_
                    sb.AppendLine(Indent & "Exit " & Identifier)
                Case SType.THROW_
                    sb.AppendLine(Indent & "Throw " & ActualParameters(0).ToString())
                Case SType.TRY_CATCH
                    sb.AppendLine(Indent & "Try")
                    Statements(0).ToStringBuilder(sb, next_indent)
                    For v_index = 1 To Variables.Count
                        sb.AppendLine(Indent & "Catch " & Variables(v_index - 1).ToString())
                        Statements(v_index).ToStringBuilder(sb, next_indent)
                    Next
                    sb.AppendLine(Indent & "End Try")
                Case SType.MODULE_
                    sb.AppendLine(Indent & IRUtils.StringOfAccessModifier(AccessModifier, " ") & "Module " & Identifier)
                    For Each st In Statements
                        st.ToStringBuilder(sb, next_indent)
                    Next
                    sb.AppendLine(Indent & "End Module")
                Case SType.NAMESPACE_
                    sb.AppendLine(Indent & IRUtils.StringOfAccessModifier(AccessModifier, " ") & "Namespace " & Identifier)
                    For Each st In Statements
                        st.ToStringBuilder(sb, next_indent)
                    Next
                    sb.AppendLine(Indent & "End Namespace")
                Case SType.ENUM_
                    sb.AppendLine(Indent & IRUtils.StringOfAccessModifier(AccessModifier, " ") & "Enum " & Identifier)
                    For Each enum_member In Variables
                        sb.Append(next_indent & enum_member.Name)
                        If enum_member.Init Is Nothing Then
                            sb.AppendLine()
                        Else
                            sb.AppendLine(" = " & enum_member.Init.ToString())
                        End If
                    Next
                    sb.AppendLine(Indent & "End Enum")
                Case SType.CLASS_
                    sb.AppendLine(Indent & IRUtils.StringOfAccessModifier(AccessModifier, " ") & "Class " & Identifier)
                    For Each st In Statements
                        st.ToStringBuilder(sb, next_indent)
                    Next
                    sb.AppendLine(Indent & "End Class")
                Case SType.FUNCTION_DECLARATION
                    ' 1. Header
                    If ReturnTypes.Count > 0 Then
                        sb.Append(Indent & "Function ")
                    Else
                        sb.Append(Indent & "Sub ")
                    End If
                    sb.Append(Identifier & "(" & IRUtils.ConcatObjects(FormalParameters, ", ") & ")")
                    If (ReturnTypes IsNot Nothing) AndAlso (ReturnTypes.Count > 0) Then
                        sb.Append(" As ")
                        sb.Append(IRUtils.ConcatObjects(ReturnTypes, ","))
                    End If
                    sb.AppendLine()

                    ' 2. Body.
                    For Each st In Statements
                        st.ToStringBuilder(sb, next_indent)
                    Next

                    ' 3. End.
                    sb.AppendLine(Indent & "End Sub")
                Case SType.OPTION_
                    sb.AppendLine(Indent & "Option " & ActualParameters(0).StringValue & " " & ActualParameters(1).StringValue)
                Case SType.IMPORTS_
                    sb.Append(Indent & "Imports ")
                    For import_index = 0 To FormalParameters.Count - 1
                        Dim im = FormalParameters(import_index)
                        If import_index > 0 Then
                            sb.Append(", ")
                        End If
                        sb.Append(im.Name)
                        If im.Type IsNot Nothing Then
                            sb.Append(" = ")
                            sb.Append(im.Type.ToString())
                        End If
                    Next
                    sb.AppendLine()
            End Select
        End Sub
    End Class
End Namespace

