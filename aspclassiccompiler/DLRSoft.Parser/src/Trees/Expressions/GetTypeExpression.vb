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
''' A parse tree for a GetType expression.
''' </summary>
Public NotInheritable Class GetTypeExpression
    Inherits Expression

    Private ReadOnly _LeftParenthesisLocation As Location
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
    ''' The target type of the GetType expression.
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

    ''' <summary>
    ''' Constructs a new parse tree for a GetType expression.
    ''' </summary>
    ''' <param name="leftParenthesisLocation">The location of the '('.</param>
    ''' <param name="target">The target type of the GetType expression.</param>
    ''' <param name="rightParenthesisLocation">The location of the ')'.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal leftParenthesisLocation As Location, ByVal target As TypeName, ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.GetTypeExpression, span)

        If target Is Nothing Then
            Throw New ArgumentNullException("target")
        End If

        SetParent(target)

        _LeftParenthesisLocation = leftParenthesisLocation
        _Target = target
        _RightParenthesisLocation = rightParenthesisLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Target)
    End Sub
End Class