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
Public NotInheritable Class QualifiedExpression
    Inherits Expression

    Private ReadOnly _Qualifier As Expression
    Private ReadOnly _DotLocation As Location
    Private ReadOnly _Name As SimpleName

    ''' <summary>
    ''' The expression qualifying the name.
    ''' </summary>
    Public ReadOnly Property Qualifier() As Expression
        Get
            Return _Qualifier
        End Get
    End Property

    ''' <summary>
    ''' The location of the '.'.
    ''' </summary>
    Public ReadOnly Property DotLocation() As Location
        Get
            Return _DotLocation
        End Get
    End Property

    ''' <summary>
    ''' The qualified name.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a qualified name expression.
    ''' </summary>
    ''' <param name="qualifier">The expression qualifying the name.</param>
    ''' <param name="dotLocation">The location of the '.'.</param>
    ''' <param name="name">The qualified name.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal qualifier As Expression, ByVal dotLocation As Location, ByVal name As SimpleName, ByVal span As Span)
        MyBase.New(TreeType.QualifiedExpression, span)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(qualifier)
        SetParent(name)

        _Qualifier = qualifier
        _DotLocation = dotLocation
        _Name = name
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Qualifier)
        AddChild(childList, Name)
    End Sub
End Class