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
''' A parse tree for an array type name.
''' </summary>
''' <remarks>
''' This tree may contain size arguments as well.
''' </remarks>
Public NotInheritable Class ArrayTypeName
    Inherits TypeName

    Private ReadOnly _ElementTypeName As TypeName
    Private ReadOnly _Rank As Integer
    Private ReadOnly _Arguments As ArgumentCollection

    ''' <summary>
    ''' The type name for the element type of the array.
    ''' </summary>
    Public ReadOnly Property ElementTypeName() As TypeName
        Get
            Return _ElementTypeName
        End Get
    End Property

    ''' <summary>
    ''' The rank of the array type name.
    ''' </summary>
    Public ReadOnly Property Rank() As Integer
        Get
            Return _Rank
        End Get
    End Property

    ''' <summary>
    ''' The arguments of the array type name, if any.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for an array type name.
    ''' </summary>
    ''' <param name="elementTypeName">The type name for the array element type.</param>
    ''' <param name="rank">The rank of the array type name.</param>
    ''' <param name="arguments">The arguments of the array type name, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal elementTypeName As TypeName, ByVal rank As Integer, ByVal arguments As ArgumentCollection, ByVal span As Span)
        MyBase.New(TreeType.ArrayType, span)

        If arguments Is Nothing Then
            Throw New ArgumentNullException("arguments")
        End If

        SetParent(elementTypeName)
        SetParent(arguments)

        _ElementTypeName = elementTypeName
        _Rank = rank
        _Arguments = arguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, ElementTypeName)
        AddChild(childList, Arguments)
    End Sub
End Class