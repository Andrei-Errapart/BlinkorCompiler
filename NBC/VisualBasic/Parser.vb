Option Compare Binary
Option Explicit On
Option Strict On

Imports System
Imports System.IO

Imports NBC.IR

Namespace VisualBasic

	Public Class Parser
		Public  Const   _EOF        As Integer =  0
		Public  Const   _ident      As Integer =  1
		Public  Const   _intliteral As Integer =  2
		Public  Const   _newline1   As Integer =  3
		Public  Const   _string     As Integer =  4
		Public  Const   maxT        As Integer = 91
		Private Const   blnT        As Boolean = True
		Private Const   blnX        As Boolean = False
		Private Const   minErrDist  As Integer =  2
		Public          scanner     As Scanner
		Public          errors      As Errors
		Public          t           As Token                ' last recognized token
		Public          la          As Token                ' lookahead token
		Private         errDist     As Integer = minErrDist
		' Output
		Public Result as List(Of Statement)
		' Helper function
		Public Shared Function ConcatExpression(ByVal expr as Expression, ByVal op as ExpressionType, ByVal expr2 as Expression) as Expression
			Dim r = new Expression With { .Type = op, .Args = new List(Of Expression) }
			r.Args.Add(expr)
			r.Args.Add(expr2)
			Return r
		End Function
		Public Shared Function ConcatUnaryExpression(ByVal op as ExpressionType, ByVal expr2 as Expression) as Expression
			Dim r = new Expression With { .Type = op, .Args = new List(Of Expression) }
			r.Args.Add(expr2)
			Return r
		End Function
		Public Function IsTemplateParams() As Boolean
			Return (la.val="(") And StrComp(scanner.Peek().val, "of", vbTextCompare)=0
		End Function
		Public Sub New(ByVal scanner As Scanner)
			Me.scanner = scanner
			errors = New Errors()
		End Sub
		Private Sub SynErr(ByVal n As Integer)
			If errDist >= minErrDist Then
				errors.SynErr(la.line, la.col, n)
			End If
			errDist = 0
		End Sub
		Public Sub SemErr(ByVal msg As String)
			If errDist >= minErrDist Then
				errors.SemErr(t.line, t.col, msg)
			End If
			errDist = 0
		End Sub
		Private Sub [Get]()
			While True
				t = la
				la = scanner.Scan()
				If la.kind <= maxT Then
					errDist += 1
					Exit While
				End If
				la = t
			End While
		End Sub
		Private Sub Expect(ByVal n As Integer)
			If la.kind = n Then
				[Get]()
			Else
				SynErr(n)
			End If
		End Sub
		Private Function StartOf(ByVal s As Integer) As Boolean
			Return blnSet(s, la.kind)
		End Function
		Private Sub ExpectWeak(ByVal n As Integer, ByVal follow As Integer)
			If la.kind = n Then
				[Get]()
			Else
				SynErr(n)
				While Not StartOf(follow)
					[Get]()
				End While
			End If
		End Sub
		Private Function WeakSeparator(ByVal n As Integer, ByVal syFol As Integer, ByVal repFol As Integer) As Boolean
			Dim kind As Integer = la.kind
			If kind = n Then
				[Get]()
				Return True
			ElseIf StartOf(repFol) Then
				Return False
			Else
				SynErr(n)
				While Not (blnSet(syFol, kind) OrElse blnSet(repFol, kind) OrElse blnSet(0, kind))
					[Get]()
					kind = la.kind
				End While
				Return StartOf(syFol)
			End If
		End Function
		Private Sub VisualBasicCompiler()
			Dim st as Statement = Nothing
			Dim am as AccessModifier = AccessModifier.NONE
				Result = new List(Of Statement)
			If StartOf(1) Then
				NonEmptyBlockOfStmt(st)
				Result.Add(VBUtils.CreateMainModule(st))
			ElseIf StartOf(2) Then
				While la.kind = 15
					OptionStmt(st)
					Result.Add(st)
				End While
				While la.kind = 20
					ImportsStmt(st)
					Result.Add(st)
				End While
				If StartOf(3) Then
					AccessModifierDeclaration(am)
				End If
				NamespaceMemberDeclaration(st, am)
				Result.Add(st)
				am = AccessModifier.NONE
				While StartOf(4)
					If StartOf(3) Then
						AccessModifierDeclaration(am)
					End If
					NamespaceMemberDeclaration(st, am)
					Result.Add(st)
					am = AccessModifier.NONE
				End While
			Else
				SynErr(92)
			End If
		End Sub
		Private Sub NonEmptyBlockOfStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.MULTIPLE_STATEMENTS, .Statements = new List(Of Statement) }
			Dim st2 as Statement = Nothing
			NonDeclaringStmt(st2)
			st.Statements.Add(st2)
			While StartOf(1)
				NonDeclaringStmt(st2)
				st.Statements.Add(st2)
			End While
		End Sub
		Private Sub OptionStmt(ByRef st As Statement)
			st = new Statement With { .Type = SType.OPTION_, .ActualParameters = new List(Of Expression) }
			Dim exp as Expression = Nothing
			Expect(15)
			Ident(exp)
			st.ActualParameters.Add(exp)
			Ident(exp)
			st.ActualParameters.Add(exp)
			newline()
		End Sub
		Private Sub ImportsStmt(ByRef st As Statement)
			st = new Statement With { .Type = SType.IMPORTS_, .FormalParameters = new List(Of Variable) }
			Dim id as String = ""
			Expect(20)
			QualifiedIdent(id)
			st.FormalParameters.Add(new Variable With { .Name = id } )
			If la.kind = 6 Then
				[Get]()
				QualifiedIdent(id)
				st.FormalParameters.Last().Type = new VariableType With { .Type = VTType.NAMED, .Name = id }
			End If
			While la.kind = 1
				QualifiedIdent(id)
				st.FormalParameters.Add(new Variable With { .Name = id } )
				If la.kind = 6 Then
					[Get]()
					QualifiedIdent(id)
					st.FormalParameters.Last().Type = new VariableType With { .Type = VTType.NAMED, .Name = id }
				End If
			End While
			newline()
		End Sub
		Private Sub AccessModifierDeclaration(ByRef am As AccessModifier)
			If la.kind = 54 Then
				[Get]()
				am = AccessModifier.PUBLIC_
			ElseIf la.kind = 55 Then
				[Get]()
				am = AccessModifier.PROTECTED_
			ElseIf la.kind = 56 Then
				[Get]()
				am = AccessModifier.FRIEND_
			ElseIf la.kind = 57 Then
				[Get]()
				am = AccessModifier.PRIVATE_
			Else
				SynErr(93)
			End If
		End Sub
		Private Sub NamespaceMemberDeclaration(ByRef st As Statement, ByVal am As AccessModifier)
			If la.kind = 12 Then
				NamespaceDeclaration(st, am)
			ElseIf la.kind = 11 Then
				ModuleDeclaration(st, am)
			ElseIf la.kind = 5 OrElse la.kind = 8 OrElse la.kind = 9 Then
				NonModuleDeclaration(st, am)
			Else
				SynErr(94)
			End If
		End Sub
		Private Sub BlockOfStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.MULTIPLE_STATEMENTS, .Statements = new List(Of Statement) }
			Dim st2 as Statement = Nothing
			While StartOf(1)
				NonDeclaringStmt(st2)
				st.Statements.Add(st2)
			End While
		End Sub
		Private Sub NonDeclaringStmt(ByRef st as Statement)
			Select Case la.kind
				Case 31
					PrintStmt(st)
				Case 32
					ForStmt(st)
				Case 10, 13
					DimStatement(st)
				Case 37
					IfStmt(st)
				Case 41
					WhileStmt(st)
				Case 46
					ThrowStmt(st)
				Case 47
					TryCatchStmt(st)
				Case 1, 2, 4, 16, 50, 65, 69
					AssignmentOrInvokeStmt(st)
				Case 44
					ExitStmt(st)
				Case 45
					ReturnStmt(st)
				Case 42
					SelectStmt(st)
				Case Else
					SynErr(95)
			End Select
		End Sub
		Private Sub PrintStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.SUBROUTINE_CALL, .Identifier = VBUtils.NameOfPrint, .ActualParameters = new List(Of Expression) }
			Dim exp as Expression = Nothing
			Expect(31)
			Expr(exp)
			st.ActualParameters.Add(exp)
			While la.kind = 14
				[Get]()
				Expr(exp)
				st.ActualParameters.Add(exp)
			End While
			newline()
		End Sub
		Private Sub ForStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.FOR_, .Variables = new List(Of Variable), .ActualParameters = new List(Of Expression) }
			Dim v as Variable = Nothing
				Dim block as Statement = Nothing
				Dim exp as Expression = Nothing
			Expect(32)
			If la.kind = 33 Then
				[Get]()
				st.Type = SType.FOR_EACH
				VariableDeclaration(v)
				st.Variables.Add(v)
				Expect(34)
				Expr(exp)
				st.ActualParameters.Add(exp)
			ElseIf la.kind = 1 Then
				VariableDeclaration(v)
				st.Variables.Add(v)
				Expect(6)
				Expr(exp)
				st.Variables.Last().Init = exp
				Expect(35)
				Expr(exp)
				st.ActualParameters.Add(exp)
			Else
				SynErr(96)
			End If
			newline()
			BlockOfStmt(block)
			st.Statements = block.Statements
			Expect(36)
			newline()
		End Sub
		Private Sub DimStatement(ByRef st As Statement)
			st = new Statement With { .Type = SType.VARIABLE_DECLARATION, .Variables = new List(Of Variable)  }
			Dim v as Variable = Nothing
				Dim vm as VariableModifier = VariableModifier.NONE
				
			If la.kind = 13 Then
				[Get]()
			ElseIf la.kind = 10 Then
				[Get]()
				vm = VariableModifier.CONST_
			Else
				SynErr(97)
			End If
			VariableDeclarationOptionalInit(v)
			v.Modifier = vm
			st.Variables.Add(v)
			While la.kind = 14
				[Get]()
				VariableDeclarationOptionalInit(v)
				v.Modifier = vm
				st.Variables.Add(v)
			End While
			newline()
		End Sub
		Private Sub IfStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.IF_, .ActualParameters = new List(Of Expression), .Statements = new List(Of Statement) }
			Dim block as Statement = Nothing
				Dim exp as Expression = Nothing
			Expect(37)
			Expr(exp)
			st.ActualParameters.Add(exp)
			Expect(38)
			newline()
			BlockOfStmt(block)
			st.Statements.Add(block)
			While la.kind = 39
				[Get]()
				Expr(exp)
				st.ActualParameters.Add(exp)
				Expect(38)
				newline()
				BlockOfStmt(block)
				st.Statements.Add(block)
			End While
			If la.kind = 40 Then
				[Get]()
				newline()
				BlockOfStmt(block)
				st.Statements.Add(block)
			End If
			Expect(7)
			Expect(37)
			newline()
		End Sub
		Private Sub WhileStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.WHILE_, .ActualParameters = new List(Of Expression), .Statements = new List(Of Statement) }
			Dim block as Statement = Nothing
				Dim exp as Expression = Nothing
			Expect(41)
			Expr(exp)
			newline()
			st.ActualParameters.Add(exp)
			BlockOfStmt(block)
			st.Statements.Add(block)
			Expect(7)
			Expect(41)
			newline()
		End Sub
		Private Sub ThrowStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.THROW_, .ActualParameters = new List(Of Expression) }
			Dim exp as Expression = Nothing
			Expect(46)
			Expr(exp)
			newline()
			st.ActualParameters.Add(exp)
		End Sub
		Private Sub TryCatchStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.TRY_CATCH, .Variables = new List(Of Variable), .Statements = new List(Of Statement) }
			Dim block as Statement = Nothing
				Dim v as Variable = Nothing
				
			Expect(47)
			newline()
			BlockOfStmt(block)
			st.Statements.Add(block)
			Expect(48)
			VariableDeclaration(v)
			st.Variables.Add(v)
			newline()
			BlockOfStmt(block)
			st.Statements.Add(block)
			While la.kind = 48
				[Get]()
				VariableDeclaration(v)
				st.Variables.Add(v)
				newline()
				BlockOfStmt(block)
				st.Statements.Add(block)
			End While
			Expect(7)
			Expect(47)
			newline()
		End Sub
		Private Sub AssignmentOrInvokeStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.INVOKE_EXPRESSION, .ActualParameters = new List(Of Expression) }
			Dim exp as Expression = Nothing
				Dim rhs as Expression = Nothing
			Expr(exp)
			st.ActualParameters.Add(exp)
			If StartOf(5) Then
				Select Case la.kind
					Case 22
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_POWER, rhs)
					Case 23
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_MULTIPLY, rhs)
					Case 24
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_DIVIDE, rhs)
					Case 25
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_INT_DIVIDE, rhs)
					Case 26
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_PLUS, rhs)
					Case 27
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_MINUS, rhs)
					Case 28
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_SHIFT_LEFT, rhs)
					Case 29
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_SHIFT_RIGHT, rhs)
					Case 30
						[Get]()
						Expr(rhs)
						VBUtils.RewriteCompositeAssignment(st, ExpressionType.BINARY_OP_CONCAT, rhs)
				End Select
			End If
			newline()
			VBUtils.CorrectForAssignment(st)
		End Sub
		Private Sub ExitStmt(ByRef st as Statement)

			Expect(44)
			st = new Statement With { .Type = SType.EXIT_, .Identifier = "" }
			If la.kind = 32 OrElse la.kind = 41 Then
				If la.kind = 41 Then
					[Get]()
					st.Identifier = "While"
				Else
					[Get]()
					st.Identifier = "For"
			End If
			End If
			newline()
		End Sub
		Private Sub ReturnStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.RETURN_, .ActualParameters = new List(Of Expression) }
			Dim exp as Expression = Nothing
			Expect(45)
			Expr(exp)
			newline()
			st.ActualParameters.Add(exp)
		End Sub
		Private Sub SelectStmt(ByRef st as Statement)
			st = new Statement With { .Type = SType.SELECT_, .ActualParameters = new List(Of Expression), .Statements = new List(Of Statement) }
			Dim block as Statement = Nothing
				Dim exp as Expression = Nothing
			Expect(42)
			Expect(43)
			Expr(exp)
			newline()
			st.ActualParameters.Add(exp)
			While la.kind = 43
				[Get]()
				If la.kind = 40 Then
					[Get]()
				ElseIf StartOf(6) Then
					Expr(exp)
					st.ActualParameters.Add(exp)
				Else
					SynErr(98)
				End If
				newline()
				BlockOfStmt(block)
				st.Statements.Add(block)
			End While
			Expect(7)
			Expect(42)
			newline()
		End Sub
		Private Sub NamespaceDeclaration(ByRef st as Statement, am as AccessModifier)
			st = new Statement With { .Type = SType.NAMESPACE_, .AccessModifier = am, .Statements = new List(Of Statement) }
			Dim qid as String = Nothing
				Dim st2 as Statement = Nothing
				Dim am2 as AccessModifier = AccessModifier.NONE
			Expect(12)
			QualifiedIdent(qid)
			st.Identifier = qid
			newline()
			While StartOf(3)
				AccessModifierDeclaration(am2)
				If la.kind = 5 OrElse la.kind = 8 OrElse la.kind = 9 Then
					NonModuleDeclaration(st2, am2)
					st.Statements.Add(st2)
				ElseIf la.kind = 11 Then
					ModuleDeclaration(st2, am2)
					st.Statements.Add(st2)
				Else
					SynErr(99)
				End If
				am2 = AccessModifier.NONE
			End While
			Expect(7)
			Expect(12)
			newline()
		End Sub
		Private Sub ModuleDeclaration(ByRef st as Statement, am as AccessModifier)
			st = new Statement With { .Type = SType.MODULE_,  .AccessModifier = am, .Statements = new List(Of Statement) }
			Dim st_sub as Statement = Nothing
				Dim am2 as AccessModifier = AccessModifier.NONE
			Expect(11)
			Expect(1)
			st.Identifier = t.val
			newline()
			While StartOf(7)
				If StartOf(3) Then
					AccessModifierDeclaration(am2)
				End If
				If StartOf(8) Then
					MethodMemberDeclaration(st_sub, am2)
					st.Statements.Add(st_sub)
				ElseIf la.kind = 10 OrElse la.kind = 13 Then
					DimStatement(st_sub)
					st.Statements.Add(st_sub)
				Else
					SynErr(100)
				End If
				am2 = AccessModifier.NONE
			End While
			Expect(7)
			Expect(11)
			newline()
		End Sub
		Private Sub NonModuleDeclaration(ByRef st As Statement, am as AccessModifier)
			If la.kind = 5 Then
				EnumDeclaration(st, am)
			ElseIf la.kind = 9 Then
				StructureDeclaration(st, am)
			ElseIf la.kind = 8 Then
				ClassDeclaration(st, am)
			Else
				SynErr(101)
			End If
		End Sub
		Private Sub EnumDeclaration(ByRef st as Statement, am as AccessModifier)
			Dim exp as Expression = Nothing
			Expect(5)
			Expect(1)
			st = new Statement With { .Type = SType.ENUM_, .AccessModifier = am, .Identifier = t.val, .Variables = new List(Of Variable) }
			newline()
			While la.kind = 3
				newline()
			End While
			Expect(1)
			st.Variables.Add(new Variable With { .Name = t.val })
			If la.kind = 6 Then
				[Get]()
				Expr(exp)
				st.Variables.Last().Init = exp
			End If
			While la.kind = 3
				newline()
			End While
			While la.kind = 1
				[Get]()
				st.Variables.Add(new Variable With { .Name = t.val })
				If la.kind = 6 Then
					[Get]()
					Expr(exp)
					st.Variables.Last().Init = exp
				End If
				While la.kind = 3
					newline()
				End While
			End While
			Expect(7)
			Expect(5)
			newline()
		End Sub
		Private Sub StructureDeclaration(ByRef st_class as Statement, class_am as AccessModifier)
			Dim am as AccessModifier = AccessModifier.NONE

			Expect(9)
			Expect(1)
			st_class = new Statement With { .Type = SType.CLASS_,  .Identifier = t.val,  .AccessModifier = class_am, .Statements = new List(Of Statement), .Variables = new List(Of Variable) }
			newline()
			While StartOf(3)
				AccessModifierDeclaration(am)
				ClassMemberDeclaration(st_class, am)
			End While
			Expect(7)
			Expect(9)
			newline()
		End Sub
		Private Sub ClassDeclaration(ByRef st_class as Statement, class_am as AccessModifier)
			Dim am as AccessModifier = AccessModifier.NONE
			Dim st as Statement = Nothing
				Dim v as Variable = Nothing
				Dim vm as VariableModifier = VariableModifier.NONE
				
			Expect(8)
			Expect(1)
			st_class = new Statement With { .Type = SType.CLASS_,  .Identifier = t.val,  .AccessModifier = class_am, .Statements = new List(Of Statement), .Variables = new List(Of Variable) }
			newline()
			While StartOf(3)
				AccessModifierDeclaration(am)
				ClassMemberDeclaration(st_class, am)
			End While
			Expect(7)
			Expect(8)
			newline()
		End Sub
		Private Sub newline()
			Expect(3)
			While la.kind = 3
				[Get]()
			End While
		End Sub
		Private Sub Expr(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim op as ExpressionType
			AndExp(exp)
			While la.kind = 88 OrElse la.kind = 89 OrElse la.kind = 90
				OrOp(op)
				AndExp(exp2)
				exp = ConcatExpression(exp, op, exp2)
			End While
		End Sub
		Private Sub ClassMemberDeclaration(ByVal st_class as Statement, ByVal am as AccessModifier)
			Dim vm as VariableModifier = VariableModifier.NONE
			Dim st as Statement = Nothing
			Dim v as Variable = Nothing

			If la.kind = 5 OrElse la.kind = 8 OrElse la.kind = 9 Then
				NonModuleDeclaration(st, am)
				st_class.Statements.Add(st)
			ElseIf StartOf(8) Then
				MethodMemberDeclaration(st, am)
				st_class.Statements.Add(st)
			ElseIf la.kind = 1 OrElse la.kind = 10 Then
				If la.kind = 10 Then
					[Get]()
					vm = VariableModifier.CONST_
				End If
				VariableDeclarationOptionalInit(v)
				newline()
				v.Modifier = vm
				st_class.Variables.Add(v)
			Else
				SynErr(102)
			End If
		End Sub
		Private Sub MethodMemberDeclaration(ByRef st as Statement, am_sub As AccessModifier)
			Dim block as Statement = Nothing
			Dim v as Variable = Nothing
				Dim pml as List(Of ProcedureModifier) = Nothing
				Dim te as VariableType = Nothing
				Dim vm as VariableModifier = VariableModifier.NONE
				
			If StartOf(9) Then
				ProcedureModifiers(pml)
			End If
			st = new Statement With { .Type = SType.FUNCTION_DECLARATION, .AccessModifier = am_sub, .Modifiers = pml, .FormalParameters = new List(Of Variable), .ReturnTypes = new List(Of VariableType)  }
			If la.kind = 49 Then
				[Get]()
				If la.kind = 50 Then
					[Get]()
					st.Identifier = "New"
					if st.Modifiers Is Nothing Then
							st.Modifiers = new List(Of ProcedureModifier)
						End If
						st.Modifiers.Add(ProcedureModifier.CONSTRUCTOR)
				ElseIf la.kind = 1 Then
					[Get]()
					st.Identifier = t.val
				Else
					SynErr(103)
				End If
				Expect(16)
				If la.kind = 1 OrElse la.kind = 52 OrElse la.kind = 53 Then
					VariableDeclarationOptionalModifier(v)
					st.FormalParameters.Add(v)
					While la.kind = 14
						[Get]()
						VariableDeclarationOptionalModifier(v)
						st.FormalParameters.Add(v)
					End While
				End If
				Expect(17)
				newline()
				BlockOfStmt(block)
				st.Statements = block.Statements
				Expect(7)
				Expect(49)
			ElseIf la.kind = 51 Then
				[Get]()
				Expect(1)
				st.Identifier = t.val
				Expect(16)
				If la.kind = 1 OrElse la.kind = 52 OrElse la.kind = 53 Then
					VariableDeclarationOptionalModifier(v)
					st.FormalParameters.Add(v)
					While la.kind = 14
						[Get]()
						VariableDeclarationOptionalModifier(v)
						st.FormalParameters.Add(v)
					End While
				End If
				Expect(17)
				Expect(18)
				TypeExpression(te)
				newline()
				st.ReturnTypes.Add(te)
				BlockOfStmt(block)
				st.Statements = block.Statements
				Expect(7)
				Expect(51)
			Else
				SynErr(104)
			End If
			newline()
		End Sub
		Private Sub VariableDeclarationOptionalInit(ByRef v as Variable)
			Dim exp As Expression = Nothing
			VariableDeclaration(v)
			If la.kind = 6 Then
				[Get]()
				Expr(exp)
				v.Init = exp
			End If
		End Sub
		Private Sub QualifiedIdent(ByRef Name As String)
			Expect(1)
			Name = t.val
			While la.kind = 21
				[Get]()
				Expect(1)
				Name = Name & "." & t.val
			End While
		End Sub
		Private Sub Ident(ByRef id as Expression)
			Expect(1)
			id = new Expression with { .Type = ExpressionType.VARIABLE_ACCESS,  .StringValue = t.val }
		End Sub
		Private Sub TypenameExpr(ByRef tn as VBTypename)
			tn = new VBTypename With { .Ranks = new List(Of List(Of Integer)) }
			Expect(1)
			tn.Name = t.val
			While la.kind = 16
				[Get]()
				tn.Ranks.Add(new List(Of Integer))
				If la.kind = 2 Then
					[Get]()
					tn.Ranks.Last().Add(Convert.ToInt32(t.val))
					While la.kind = 14
						[Get]()
						Expect(2)
						tn.Ranks.Last().Add(Convert.ToInt32(t.val))
					End While
				End If
				Expect(17)
				If tn.Ranks.Last().Count = 0 Then
				tn.Ranks.Last().Add(-1)
				End If

			End While
		End Sub
		Private Sub VariableDeclaration(ByRef v as Variable)
			Dim tn as VBTypename = Nothing
			Dim te as VariableType = Nothing
			TypenameExpr(tn)
			v = Variable.CreateFromTypename(tn)
			If la.kind = 18 Then
				[Get]()
				TypeExpression(te)
				v.SetMostNestedType(te)
			End If
		End Sub
		Private Sub TypeExpression(ByRef te as VariableType)
			Dim qid as String = Nothing
			Dim t2 as VariableType = Nothing
			QualifiedIdent(qid)
			te = new VariableType With { .Type=VTType.NAMED, .Name=qid }
			If IsTemplateParams() Then
				Expect(16)
				Expect(19)
				TypeExpression(t2)
				te.NestedType = new List(Of VariableType)
				te.NestedType.Add(t2)
				While la.kind = 14
					[Get]()
					TypeExpression(t2)
					te.NestedType.Add(t2)
				End While
				Expect(17)
			End If
		End Sub
		Private Sub ProcedureModifiers(ByRef pml as List(Of ProcedureModifier))
			pml = new List(Of ProcedureModifier)
			Dim pm as ProcedureModifier
			ProcedureModifierDeclaration(pm)
			pml.Add(pm)
			While StartOf(9)
				ProcedureModifierDeclaration(pm)
				pml.Add(pm)
			End While
		End Sub
		Private Sub VariableDeclarationOptionalModifier(ByRef v as Variable)
			Dim vm as VariableModifier = VariableModifier.NONE
			If la.kind = 52 OrElse la.kind = 53 Then
				If la.kind = 52 Then
					[Get]()
					vm = VariableModifier.BYVAL_
				Else
					[Get]()
					vm = VariableModifier.BYREF_
			End If
			End If
			VariableDeclaration(v)
			v.Modifier = vm
		End Sub
		Private Sub ProcedureModifierDeclaration(ByRef pm as ProcedureModifier)
			Select Case la.kind
				Case 58
					[Get]()
					pm = ProcedureModifier.SHADOWS_
				Case 59
					[Get]()
					pm = ProcedureModifier.SHARED_
				Case 60
					[Get]()
					pm = ProcedureModifier.OVERRIDABLE_
				Case 61
					[Get]()
					pm = ProcedureModifier.NOTOVERRIDABLE_
				Case 62
					[Get]()
					pm = ProcedureModifier.MUSTOVERRIDE_
				Case 63
					[Get]()
					pm = ProcedureModifier.OVERRIDES_
				Case 64
					[Get]()
					pm = ProcedureModifier.OVERLOADS_
				Case Else
					SynErr(105)
			End Select
		End Sub
		Private Sub Number(ByRef num as Expression)
			Expect(2)
			num = new Expression with { .Type = ExpressionType.INTEGER_, .IntegerValue = Convert.ToInt32(t.val) }
		End Sub
		Private Sub AndExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim op as ExpressionType
			NotExp(exp)
			While la.kind = 86 OrElse la.kind = 87
				AndOp(op)
				NotExp(exp2)
				exp = ConcatExpression(exp, op, exp2)
			End While
		End Sub
		Private Sub OrOp(ByRef op As ExpressionType)
			If la.kind = 88 Then
				[Get]()
				op = ExpressionType.BINARY_OP_OR
			ElseIf la.kind = 89 Then
				[Get]()
				op = ExpressionType.BINARY_OP_ORELSE
			ElseIf la.kind = 90 Then
				[Get]()
				op = ExpressionType.BINARY_OP_XOR
			Else
				SynErr(106)
			End If
		End Sub
		Private Sub NotExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			If StartOf(10) Then
				CompareExp(exp)
			ElseIf la.kind = 65 Then
				[Get]()
				CompareExp(exp2)
				exp = ConcatUnaryExpression(ExpressionType.UNARY_OP_NOT, exp2)
			Else
				SynErr(107)
			End If
		End Sub
		Private Sub AndOp(ByRef op As ExpressionType)
			If la.kind = 86 Then
				[Get]()
				op = ExpressionType.BINARY_OP_AND
			ElseIf la.kind = 87 Then
				[Get]()
				op = ExpressionType.BINARY_OP_ANDALSO
			Else
				SynErr(108)
			End If
		End Sub
		Private Sub CompareExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim op as ExpressionType
			ShiftExp(exp)
			While StartOf(11)
				CompareOp(op)
				ShiftExp(exp2)
				exp = ConcatExpression(exp, op, exp2)
			End While
		End Sub
		Private Sub ShiftExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim op as ExpressionType
			ConcatExp(exp)
			While la.kind = 84 OrElse la.kind = 85
				ShiftOp(op)
				ConcatExp(exp2)
				exp = ConcatExpression(exp, op, exp2)
			End While
		End Sub
		Private Sub CompareOp(ByRef op As ExpressionType)
			Select Case la.kind
				Case 6
					[Get]()
					op = ExpressionType.BINARY_OP_EQUALS
				Case 77
					[Get]()
					op = ExpressionType.BINARY_OP_NOT_EQUALS
				Case 78
					[Get]()
					op = ExpressionType.BINARY_OP_LESSER
				Case 79
					[Get]()
					op = ExpressionType.BINARY_OP_LESSER_OR_EQUAL
				Case 80
					[Get]()
					op = ExpressionType.BINARY_OP_GREATER
				Case 81
					[Get]()
					op = ExpressionType.BINARY_OP_GREATER_OR_EQUAL
				Case 82
					[Get]()
					op = ExpressionType.BINARY_OP_IS
				Case 83
					[Get]()
					op = ExpressionType.BINARY_OP_ISNOT
				Case Else
					SynErr(109)
			End Select
		End Sub
		Private Sub ConcatExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			AddExp(exp)
			While la.kind = 66
				[Get]()
				AddExp(exp2)
				exp = ConcatExpression(exp, ExpressionType.BINARY_OP_CONCAT, exp2)
			End While
		End Sub
		Private Sub ShiftOp(ByRef op As ExpressionType)
			If la.kind = 84 Then
				[Get]()
				op = ExpressionType.BINARY_OP_SHIFT_LEFT
			ElseIf la.kind = 85 Then
				[Get]()
				op = ExpressionType.BINARY_OP_SHIFT_RIGHT
			Else
				SynErr(110)
			End If
		End Sub
		Private Sub AddExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim op as ExpressionType
			ModulusExp(exp)
			While la.kind = 69 OrElse la.kind = 74
				AddOp(op)
				ModulusExp(exp2)
				exp = ConcatExpression(exp, op, exp2)
			End While
		End Sub
		Private Sub ModulusExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			IntDivExp(exp)
			While la.kind = 67
				[Get]()
				IntDivExp(exp2)
				exp = ConcatExpression(exp, ExpressionType.BINARY_OP_MODULUS, exp2)
			End While
		End Sub
		Private Sub AddOp(ByRef op As ExpressionType)
			If la.kind = 74 Then
				[Get]()
				op = ExpressionType.BINARY_OP_PLUS
			ElseIf la.kind = 69 Then
				[Get]()
				op = ExpressionType.BINARY_OP_MINUS
			Else
				SynErr(111)
			End If
		End Sub
		Private Sub IntDivExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			MultExp(exp)
			While la.kind = 68
				[Get]()
				MultExp(exp2)
				exp = ConcatExpression(exp, ExpressionType.BINARY_OP_INT_DIVIDE, exp2)
			End While
		End Sub
		Private Sub MultExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim op as ExpressionType
				
			NegateExp(exp)
			While la.kind = 75 OrElse la.kind = 76
				MulOp(op)
				NegateExp(exp2)
				exp = ConcatExpression(exp, op, exp2)
			End While
		End Sub
		Private Sub NegateExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			If StartOf(12) Then
				PowerExp(exp)
			ElseIf la.kind = 69 Then
				[Get]()
				PowerExp(exp2)
				exp = ConcatUnaryExpression(ExpressionType.UNARY_OP_MINUS, exp2)
			Else
				SynErr(112)
			End If
		End Sub
		Private Sub MulOp(ByRef op As ExpressionType)
			If la.kind = 75 Then
				[Get]()
				op = ExpressionType.BINARY_OP_MULTIPLY
			ElseIf la.kind = 76 Then
				[Get]()
				op = ExpressionType.BINARY_OP_DIVIDE
			Else
				SynErr(113)
			End If
		End Sub
		Private Sub PowerExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			FieldOrInvokeExp(exp)
			While la.kind = 70
				[Get]()
				FieldOrInvokeExp(exp2)
				exp = ConcatExpression(exp, ExpressionType.BINARY_OP_POWER, exp2)
			End While
		End Sub
		Private Sub FieldOrInvokeExp(ByRef exp as Expression)
			Dim exp2 As Expression = Nothing
			Value(exp)
			While la.kind = 16 OrElse la.kind = 21
				If la.kind = 16 Then
					[Get]()
					exp = new Expression With { .Type = ExpressionType.FUNCTION_CALL, .ExpressionValue = exp, .Args = new List(Of Expression) }
					If StartOf(6) Then
						Expr(exp2)
						exp.Args.Add(exp2)
						While la.kind = 14
							[Get]()
							Expr(exp2)
							exp.Args.Add(exp2)
						End While
					End If
					Expect(17)
				Else
					[Get]()
					Expect(1)
					exp = new Expression With { .Type = ExpressionType.MEMBER_ACCESS, .ExpressionValue = exp, .StringValue = t.val }
			End If
			End While
		End Sub
		Private Sub Value(ByRef exp as Expression)
			If la.kind = 1 Then
				Ident(exp)
			ElseIf la.kind = 2 Then
				Number(exp)
			ElseIf la.kind = 4 Then
				[Get]()
				exp = new Expression With { .Type = ExpressionType.STRING_,  .StringValue = t.val.Substring(1, t.val.Length-2) }
			ElseIf la.kind = 16 Then
				[Get]()
				Expr(exp)
				Expect(17)
			ElseIf la.kind = 50 Then
				NewExp(exp)
			Else
				SynErr(114)
			End If
		End Sub
		Private Sub NewExp(ByRef exp as Expression)
			Dim exp2 as Expression = Nothing
			Dim ival as String = Nothing
			 Dim te as VariableType = Nothing
			Expect(50)
			TypeExpression(te)
			exp = new Expression With { .Type = ExpressionType.NEW_, .NewType = te, .Args = new List(Of Expression) }
			If la.kind = 16 Then
				[Get]()
				If StartOf(6) Then
					Expr(exp2)
					exp.Args.Add(exp2)
					While la.kind = 14
						[Get]()
						Expr(exp2)
						exp.Args.Add(exp2)
					End While
				End If
				Expect(17)
			End If
			If la.kind = 71 Then
				[Get]()
				Expect(72)
				exp.Initializers = new List(Of KeyValuePair(Of String, Expression))
				Expect(21)
				SimpleIdent(ival)
				Expect(6)
				Expr(exp2)
				exp.Initializers.Add(new KeyValuePair(Of String, Expression)(ival, exp2))
				While la.kind = 14
					[Get]()
					Expect(21)
					SimpleIdent(ival)
					Expect(6)
					Expr(exp2)
					exp.Initializers.Add(new KeyValuePair(Of String, Expression)(ival, exp2))
				End While
				Expect(73)
			End If
		End Sub
		Private Sub SimpleIdent(ByRef Name As String)
			Expect(1)
			Name = t.val
		End Sub
		Public Sub Parse()
			la = New Token()
			la.val = ""
			[Get]()
			VisualBasicCompiler()
			Expect(0)
		End Sub
		Private Shared ReadOnly blnSet(,) As Boolean = { _
			{blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnT,blnT,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnT,blnX, blnX,blnT,blnX,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnT, blnT,blnX,blnX,blnX, blnX,blnT,blnX,blnX, blnX,blnT,blnT,blnX, blnT,blnT,blnT,blnT, blnX,blnX,blnT,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnX, blnX,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnX, blnT,blnT,blnX,blnT, blnT,blnX,blnX,blnT, blnX,blnX,blnX,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnT, blnT,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnT, blnT,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnX, blnT,blnT,blnX,blnT, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnT, blnT,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnT, blnT,blnT,blnT,blnT, blnT,blnT,blnT,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnT,blnT,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnX, blnX,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnX, blnX,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnT, blnX,blnX,blnT,blnT, blnT,blnT,blnT,blnT, blnT,blnT,blnT,blnT, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnT, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnT, blnT,blnT,blnT,blnT, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnT, blnT,blnT,blnT,blnT, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnT,blnT,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnT,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnT,blnT,blnT, blnT,blnT,blnT,blnT, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX}, _
			{blnX,blnT,blnT,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnT,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnT,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX,blnX,blnX,blnX, blnX} _
		}
	End Class

	Public Class Errors
		Public count        As Integer                                 ' number of errors detected
		Public errorStream  As TextWriter = Console.Out                ' error messages go to this stream
		Public errMsgFormat As String     = "-- line {0} col {1}: {2}" ' 0=line, 1=column, 2=text
		Public Sub SynErr(ByVal line As Integer, ByVal col As Integer, ByVal n As Integer)
			Dim s As String
			Select Case n
				Case    0 : s = "EOF expected"
				Case    1 : s = "ident expected"
				Case    2 : s = "intliteral expected"
				Case    3 : s = "newline1 expected"
				Case    4 : s = "string expected"
				Case    5 : s = """enum"" expected"
				Case    6 : s = """="" expected"
				Case    7 : s = """end"" expected"
				Case    8 : s = """class"" expected"
				Case    9 : s = """structure"" expected"
				Case   10 : s = """const"" expected"
				Case   11 : s = """module"" expected"
				Case   12 : s = """namespace"" expected"
				Case   13 : s = """dim"" expected"
				Case   14 : s = ""","" expected"
				Case   15 : s = """option"" expected"
				Case   16 : s = """("" expected"
				Case   17 : s = """)"" expected"
				Case   18 : s = """as"" expected"
				Case   19 : s = """of"" expected"
				Case   20 : s = """imports"" expected"
				Case   21 : s = """."" expected"
				Case   22 : s = """^="" expected"
				Case   23 : s = """*="" expected"
				Case   24 : s = """/="" expected"
				Case   25 : s = """\\\\="" expected"
				Case   26 : s = """+="" expected"
				Case   27 : s = """-="" expected"
				Case   28 : s = """<<="" expected"
				Case   29 : s = """>>="" expected"
				Case   30 : s = """&="" expected"
				Case   31 : s = """print"" expected"
				Case   32 : s = """for"" expected"
				Case   33 : s = """each"" expected"
				Case   34 : s = """in"" expected"
				Case   35 : s = """to"" expected"
				Case   36 : s = """next"" expected"
				Case   37 : s = """if"" expected"
				Case   38 : s = """then"" expected"
				Case   39 : s = """elseif"" expected"
				Case   40 : s = """else"" expected"
				Case   41 : s = """while"" expected"
				Case   42 : s = """select"" expected"
				Case   43 : s = """case"" expected"
				Case   44 : s = """exit"" expected"
				Case   45 : s = """return"" expected"
				Case   46 : s = """throw"" expected"
				Case   47 : s = """try"" expected"
				Case   48 : s = """catch"" expected"
				Case   49 : s = """sub"" expected"
				Case   50 : s = """new"" expected"
				Case   51 : s = """function"" expected"
				Case   52 : s = """byval"" expected"
				Case   53 : s = """byref"" expected"
				Case   54 : s = """public"" expected"
				Case   55 : s = """protected"" expected"
				Case   56 : s = """friend"" expected"
				Case   57 : s = """private"" expected"
				Case   58 : s = """shadows"" expected"
				Case   59 : s = """shared"" expected"
				Case   60 : s = """overridable"" expected"
				Case   61 : s = """notoverridable"" expected"
				Case   62 : s = """mustoverride"" expected"
				Case   63 : s = """overrides"" expected"
				Case   64 : s = """overloads"" expected"
				Case   65 : s = """not"" expected"
				Case   66 : s = """&"" expected"
				Case   67 : s = """mod"" expected"
				Case   68 : s = """\\\\"" expected"
				Case   69 : s = """-"" expected"
				Case   70 : s = """^"" expected"
				Case   71 : s = """with"" expected"
				Case   72 : s = """{"" expected"
				Case   73 : s = """}"" expected"
				Case   74 : s = """+"" expected"
				Case   75 : s = """*"" expected"
				Case   76 : s = """/"" expected"
				Case   77 : s = """<>"" expected"
				Case   78 : s = """<"" expected"
				Case   79 : s = """<="" expected"
				Case   80 : s = """>"" expected"
				Case   81 : s = """>="" expected"
				Case   82 : s = """is"" expected"
				Case   83 : s = """isnot"" expected"
				Case   84 : s = """<<"" expected"
				Case   85 : s = """>>"" expected"
				Case   86 : s = """and"" expected"
				Case   87 : s = """andalso"" expected"
				Case   88 : s = """or"" expected"
				Case   89 : s = """orelse"" expected"
				Case   90 : s = """xor"" expected"
				Case   91 : s = "??? expected"
				Case   92 : s = "invalid VisualBasicCompiler"
				Case   93 : s = "invalid AccessModifierDeclaration"
				Case   94 : s = "invalid NamespaceMemberDeclaration"
				Case   95 : s = "invalid NonDeclaringStmt"
				Case   96 : s = "invalid ForStmt"
				Case   97 : s = "invalid DimStatement"
				Case   98 : s = "invalid SelectStmt"
				Case   99 : s = "invalid NamespaceDeclaration"
				Case  100 : s = "invalid ModuleDeclaration"
				Case  101 : s = "invalid NonModuleDeclaration"
				Case  102 : s = "invalid ClassMemberDeclaration"
				Case  103 : s = "invalid MethodMemberDeclaration"
				Case  104 : s = "invalid MethodMemberDeclaration"
				Case  105 : s = "invalid ProcedureModifierDeclaration"
				Case  106 : s = "invalid OrOp"
				Case  107 : s = "invalid NotExp"
				Case  108 : s = "invalid AndOp"
				Case  109 : s = "invalid CompareOp"
				Case  110 : s = "invalid ShiftOp"
				Case  111 : s = "invalid AddOp"
				Case  112 : s = "invalid NegateExp"
				Case  113 : s = "invalid MulOp"
				Case  114 : s = "invalid Value"
				Case Else : s = "error " & n
			End Select
			errorStream.WriteLine(errMsgFormat, line, col, s)
			count += 1
		End Sub
		Public Sub SemErr(ByVal line As Integer, ByVal col As Integer, ByVal s As String)
			errorStream.WriteLine(errMsgFormat, line, col, s)
			count += 1
		End Sub
		Public Sub SemErr(ByVal s As String)
			errorStream.WriteLine(s)
			count += 1
		End Sub
		Public Sub Warning(ByVal line As Integer, ByVal col As Integer, ByVal s As String)
			errorStream.WriteLine(errMsgFormat, line, col, s)
		End Sub
		Public Sub Warning(ByVal s As String)
			errorStream.WriteLine(s)
		End Sub
	End Class

	Public Class FatalError
		Inherits Exception
		Public Sub New(ByVal m As String)
			MyBase.New(m)
		End Sub
	End Class

End Namespace
