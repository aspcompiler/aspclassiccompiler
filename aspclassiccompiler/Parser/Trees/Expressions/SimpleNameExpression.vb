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
''' A parse tree for a simple name expression.
''' </summary>
Public NotInheritable Class SimpleNameExpression
    Inherits Expression

    Private ReadOnly _Name As SimpleName

    ''' <summary>
    ''' The name.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a simple name expression.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal name As SimpleName, ByVal span As Span)
        MyBase.New(TreeType.SimpleNameExpression, span)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)

        _Name = name
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
    End Sub
End Class