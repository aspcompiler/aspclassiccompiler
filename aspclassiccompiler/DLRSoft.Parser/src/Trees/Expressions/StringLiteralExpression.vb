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
''' A parse tree for a string literal expression.
''' </summary>
Public NotInheritable Class StringLiteralExpression
    Inherits LiteralExpression

    Private ReadOnly _Literal As String

    ''' <summary>
    ''' The literal value.
    ''' </summary>
    Public ReadOnly Property Literal() As String
        Get
            Return _Literal
        End Get
    End Property

    'LC
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _Literal
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new string literal expression parse tree.
    ''' </summary>
    ''' <param name="literal">The literal value.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal literal As String, ByVal span As Span)
        MyBase.New(TreeType.StringLiteralExpression, span)

        If literal Is Nothing Then
            Throw New ArgumentNullException("literal")
        End If

        _Literal = literal
    End Sub
End Class