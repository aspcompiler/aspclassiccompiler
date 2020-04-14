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
''' A parse tree for a CType or DirectCast expression.
''' </summary>
Public MustInherit Class CastTypeExpression
    Inherits UnaryExpression

    Private ReadOnly _LeftParenthesisLocation As Location
    Private ReadOnly _CommaLocation As Location
    Private ReadOnly _Target As TypeName
    Private ReadOnly _RightParenthesisLocation As Location

    ''' <summary>
    ''' The location of the '('.
    ''' </summary>
    Public ReadOnly Property LeftParenthesisLocation() As Location
        Get
            Return _LeftParenthesisLocation
        End Get
    End Property

    ''' <summary>
    ''' The location of the ','.
    ''' </summary>
    Public ReadOnly Property CommaLocation() As Location
        Get
            Return _CommaLocation
        End Get
    End Property

    ''' <summary>
    ''' The target type for the operand.
    ''' </summary>
    Public ReadOnly Property Target() As TypeName
        Get
            Return _Target
        End Get
    End Property

    ''' <summary>
    ''' The location of the ')'.
    ''' </summary>
    Public ReadOnly Property RightParenthesisLocation() As Location
        Get
            Return _RightParenthesisLocation
        End Get
    End Property

    Protected Sub New(ByVal type As TreeType, ByVal leftParenthesisLocation As Location, ByVal operand As Expression, ByVal commaLocation As Location, ByVal target As TypeName, ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(type, operand, span)

        Debug.Assert(type = TreeType.CTypeExpression OrElse type = TreeType.DirectCastExpression OrElse type = TreeType.TryCastExpression)

        If target Is Nothing Then
            Throw New ArgumentNullException("target")
        End If

        SetParent(target)

        _Target = target
        _LeftParenthesisLocation = leftParenthesisLocation
        _CommaLocation = commaLocation
        _RightParenthesisLocation = rightParenthesisLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, Target)
    End Sub
End Class