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
''' A parse tree for a ReDim statement.
''' </summary>
Public NotInheritable Class ReDimStatement
    Inherits Statement

    Private ReadOnly _PreserveLocation As Location
    Private ReadOnly _Variables As ExpressionCollection

    ''' <summary>
    ''' The location of the 'Preserve', if any.
    ''' </summary>
    Public ReadOnly Property PreserveLocation() As Location
        Get
            Return _PreserveLocation
        End Get
    End Property

    ''' <summary>
    ''' The variables to redimension (includes bounds).
    ''' </summary>
    Public ReadOnly Property Variables() As ExpressionCollection
        Get
            Return _Variables
        End Get
    End Property

    ''' <summary>
    ''' Whether the statement included a Preserve keyword.
    ''' </summary>
    Public ReadOnly Property IsPreserve() As Boolean
        Get
            Return PreserveLocation.IsValid
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a ReDim statement.
    ''' </summary>
    ''' <param name="preserveLocation">The location of the 'Preserve', if any.</param>
    ''' <param name="variables">The variables to redimension (includes bounds).</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal preserveLocation As Location, ByVal variables As ExpressionCollection, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.ReDimStatement, span, comments)

        If variables Is Nothing Then
            Throw New ArgumentNullException("variables")
        End If

        SetParent(variables)
        _PreserveLocation = preserveLocation
        _Variables = variables
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Variables)
    End Sub
End Class