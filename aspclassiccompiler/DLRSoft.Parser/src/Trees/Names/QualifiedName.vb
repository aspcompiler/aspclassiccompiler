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
''' A parse tree for a qualified name (e.g. 'foo.bar').
''' </summary>
Public NotInheritable Class QualifiedName
    Inherits Name

    Private ReadOnly _Qualifier As Name
    Private ReadOnly _DotLocation As Location
    Private ReadOnly _Name As SimpleName

    ''' <summary>
    ''' The qualifier on the left-hand side of the dot.
    ''' </summary>
    Public ReadOnly Property Qualifier() As Name
        Get
            Return _Qualifier
        End Get
    End Property

    ''' <summary>
    ''' The location of the dot.
    ''' </summary>
    Public ReadOnly Property DotLocation() As Location
        Get
            Return _DotLocation
        End Get
    End Property

    ''' <summary>
    ''' The name on the right-hand side of the dot.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a qualified name.
    ''' </summary>
    ''' <param name="qualifier">The qualifier on the left-hand side of the dot.</param>
    ''' <param name="dotLocation">The location of the dot.</param>
    ''' <param name="name">The name on the right-hand side of the dot.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal qualifier As Name, ByVal dotLocation As Location, ByVal name As SimpleName, ByVal span As Span)
        MyBase.New(TreeType.QualifiedName, span)

        If qualifier Is Nothing Then
            Throw New ArgumentNullException("qualifier")
        End If

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