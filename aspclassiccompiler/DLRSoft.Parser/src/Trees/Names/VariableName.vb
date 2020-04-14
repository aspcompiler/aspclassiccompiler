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
''' A parse tree to represent a variable name.
''' </summary>
''' <remarks>
''' A variable name can have an array modifier after it (e.g. 'x(10) As Integer').
''' </remarks>
Public NotInheritable Class VariableName
    Inherits Name

    Private ReadOnly _Name As SimpleName
    Private ReadOnly _ArrayType As ArrayTypeName

    ''' <summary>
    ''' The name.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The array modifier, if any.
    ''' </summary>
    Public ReadOnly Property ArrayType() As ArrayTypeName
        Get
            Return _ArrayType
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new variable name parse tree.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="arrayType">The array modifier, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal name As SimpleName, ByVal arrayType As ArrayTypeName, ByVal span As Span)
        MyBase.New(TreeType.VariableName, span)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(arrayType)

        _Name = name
        _ArrayType = arrayType
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
        AddChild(childList, ArrayType)
    End Sub
End Class