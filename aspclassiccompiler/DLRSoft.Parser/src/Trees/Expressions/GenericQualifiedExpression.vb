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
''' A parse tree for a qualified name expression.
''' </summary>
Public NotInheritable Class GenericQualifiedExpression
    Inherits Expression

    Private ReadOnly _Base As Expression
    Private ReadOnly _TypeArguments As TypeArgumentCollection

    ''' <summary>
    ''' The base expression.
    ''' </summary>
    Public ReadOnly Property Base() As Expression
        Get
            Return _Base
        End Get
    End Property

    ''' <summary>
    ''' The qualifying type arguments.
    ''' </summary>
    Public ReadOnly Property TypeArguments() As TypeArgumentCollection
        Get
            Return _TypeArguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a generic qualified expression.
    ''' </summary>
    ''' <param name="base">The base expression.</param>
    ''' <param name="typeArguments">The qualifying type arguments.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal base As Expression, ByVal typeArguments As TypeArgumentCollection, ByVal span As Span)
        MyBase.New(TreeType.GenericQualifiedExpression, span)

        If base Is Nothing Then
            Throw New ArgumentNullException("base")
        End If

        If typeArguments Is Nothing Then
            Throw New ArgumentNullException("typeArguments")
        End If

        SetParent(base)
        SetParent(typeArguments)

        _Base = base
        _TypeArguments = typeArguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Base)
        AddChild(childList, TypeArguments)
    End Sub
End Class