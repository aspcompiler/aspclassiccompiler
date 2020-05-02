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
''' A parse tree for a TypeOf ... Is expression.
''' </summary>
Public NotInheritable Class TypeOfExpression
    Inherits UnaryExpression

    Private ReadOnly _IsLocation As Location
    Private ReadOnly _Target As TypeName

    ''' <summary>
    ''' The location of the 'Is'.
    ''' </summary>
    Public ReadOnly Property IsLocation() As Location
        Get
            Return _IsLocation
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
    ''' Constructs a new parse tree for a TypeOf ... Is expression.
    ''' </summary>
    ''' <param name="operand">The target value.</param>
    ''' <param name="isLocation">The location of the 'Is'.</param>
    ''' <param name="target">The target type to check against.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal operand As Expression, ByVal isLocation As Location, ByVal target As TypeName, ByVal span As Span)
        MyBase.New(TreeType.TypeOfExpression, operand, span)

        If target Is Nothing Then
            Throw New ArgumentNullException("target")
        End If

        SetParent(target)

        _Target = target
        _IsLocation = isLocation
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)
        AddChild(childList, Target)
    End Sub
End Class