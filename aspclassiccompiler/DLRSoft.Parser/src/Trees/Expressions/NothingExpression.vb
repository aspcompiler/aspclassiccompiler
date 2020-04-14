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
''' A parse tree for Nothing.
''' </summary>
Public NotInheritable Class NothingExpression
    Inherits Expression

    Public Overrides ReadOnly Property IsConstant() As Boolean
        Get
            Return True
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for Nothing.
    ''' </summary>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal span As Span)
        MyBase.New(TreeType.NothingExpression, span)
    End Sub
End Class