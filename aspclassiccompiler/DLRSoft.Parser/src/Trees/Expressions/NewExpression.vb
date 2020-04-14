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
''' A parse tree for a New expression.
''' </summary>
Public NotInheritable Class NewExpression
    Inherits Expression

    Private ReadOnly _Target As TypeName
    Private ReadOnly _Arguments As ArgumentCollection

    ''' <summary>
    ''' The target type to create.
    ''' </summary>
    Public ReadOnly Property Target() As TypeName
        Get
            Return _Target
        End Get
    End Property

    ''' <summary>
    ''' The arguments to the constructor.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a New expression.
    ''' </summary>
    ''' <param name="target">The target type to create.</param>
    ''' <param name="arguments">The arguments to the constructor.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal target As TypeName, ByVal arguments As ArgumentCollection, ByVal span As Span)
        MyBase.New(TreeType.NewExpression, span)

        If target Is Nothing Then
            Throw New ArgumentNullException("target")
        End If

        SetParent(target)
        SetParent(arguments)

        _Target = target
        _Arguments = arguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Target)
        AddChild(childList, Arguments)
    End Sub
End Class