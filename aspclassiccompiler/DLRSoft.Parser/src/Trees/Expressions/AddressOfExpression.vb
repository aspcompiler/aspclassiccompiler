'
' Visual Basic .NET Parser
'
' Copyright (C) 2005, Microsoft Corporation. All rights reserved.
'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
' MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'

''' <summary>
''' A parse tree for an AddressOf expression.
''' </summary>
Public NotInheritable Class AddressOfExpression
    Inherits UnaryExpression

    ''' <summary>
    ''' Constructs a new AddressOf expression parse tree.
    ''' </summary>
    ''' <param name="operand">The operand of AddressOf.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal operand As Expression, ByVal span As Span)
        MyBase.New(TreeType.AddressOfExpression, operand, span)
    End Sub
End Class