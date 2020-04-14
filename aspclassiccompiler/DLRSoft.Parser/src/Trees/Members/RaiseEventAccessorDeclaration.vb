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
''' A parse tree for a RaiseEvent property accessor.
''' </summary>
Public NotInheritable Class RaiseEventAccessorDeclaration
    Inherits ModifiedDeclaration

    Private ReadOnly _RaiseEventLocation As Location
    Private ReadOnly _Parameters As ParameterCollection
    Private ReadOnly _Statements As StatementCollection
    Private ReadOnly _EndStatement As EndBlockDeclaration

    ''' <summary>
    ''' The location of the 'RaiseEvent'.
    ''' </summary>
    Public ReadOnly Property RaiseEventLocation() As Location
        Get
            Return _RaiseEventLocation
        End Get
    End Property

    ''' <summary>
    ''' The accessor's parameters.
    ''' </summary>
    Public ReadOnly Property Parameters() As ParameterCollection
        Get
            Return _Parameters
        End Get
    End Property

    ''' <summary>
    ''' The statements in the accessor.
    ''' </summary>
    Public ReadOnly Property Statements() As StatementCollection
        Get
            Return _Statements
        End Get
    End Property

    ''' <summary>
    ''' The End declaration for the accessor.
    ''' </summary>
    Public ReadOnly Property EndStatement() As EndBlockDeclaration
        Get
            Return _EndStatement
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new parse tree for a property accessor.
    ''' </summary>
    ''' <param name="attributes">The attributes for the parse tree.</param>
    ''' <param name="raiseEventLocation">The location of the 'RaiseEvent'.</param>
    ''' <param name="parameters">The parameters of the declaration.</param>
    ''' <param name="statements">The statements in the declaration.</param>
    ''' <param name="endStatement">The end block declaration, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    ''' <param name="comments">The comments for the parse tree.</param>
    Public Sub New(ByVal attributes As AttributeBlockCollection, ByVal raiseEventLocation As Location, ByVal parameters As ParameterCollection, ByVal statements As StatementCollection, ByVal endStatement As EndBlockDeclaration, ByVal span As Span, ByVal comments As IList(Of Comment))
        MyBase.New(TreeType.RaiseEventAccessorDeclaration, attributes, Nothing, span, comments)

        SetParent(parameters)
        SetParent(statements)
        SetParent(endStatement)

        _Parameters = parameters
        _RaiseEventLocation = raiseEventLocation
        _Statements = statements
        _EndStatement = endStatement
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        MyBase.GetChildTrees(childList)

        AddChild(childList, Parameters)
        AddChild(childList, Statements)
        AddChild(childList, EndStatement)
    End Sub
End Class