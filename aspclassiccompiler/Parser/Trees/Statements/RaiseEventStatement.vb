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
''' A parse tree for a RaiseEvent statement.
''' </summary>
Public NotInheritable Class RaiseEventStatement
    Inherits Statement

    Private ReadOnly _Name As SimpleName
    Private ReadOnly _Arguments As ArgumentCollection

    ''' <summary>
    ''' The name of the event to raise.
    ''' </summary>
    Public ReadOnly Property Name() As SimpleName
        Get
            Return _Name
        End Get
    End Property

    ''' <summary>
    ''' The arguments to the event.
    ''' </summary>
    Public ReadOnly Property Arguments() As ArgumentCollection
        Get
            Return _Arguments
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a RaiseEvents statement.
    ''' </summary>
    ''' <param name="name">The name of the event to raise.</param>
    ''' <param name="arguments">The arguments to the event.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal name As SimpleName, ByVal arguments As ArgumentCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.RaiseEventStatement, span, comments)

        If name Is Nothing Then
            Throw New ArgumentNullException("name")
        End If

        SetParent(name)
        SetParent(arguments)

        _Name = name
        _Arguments = arguments
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Name)
        AddChild(childList, Arguments)
    End Sub
End Class