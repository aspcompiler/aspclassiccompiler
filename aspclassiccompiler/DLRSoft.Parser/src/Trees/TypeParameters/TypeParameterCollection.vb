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
''' A collection of type parameters.
''' </summary>
Public NotInheritable Class TypeParameterCollection
    Inherits CommaDelimitedTreeCollection(Of TypeParameter)

    Private ReadOnly _OfLocation As Location
    Private ReadOnly _RightParenthesisLocation As Location

    ''' <summary>
    ''' The location of the 'Of'.
    ''' </summary>
    Public ReadOnly Property OfLocation() As Location
        Get
            Return _OfLocation
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
    ''' Constructs a new collection of type parameters.
    ''' </summary>
    ''' <param name="ofLocation">The location of the 'Of'.</param>
    ''' <param name="parameters">The type parameters in the collection</param>
    ''' <param name="commaLocations">The locations of the commas.</param>
    ''' <param name="rightParenthesisLocation">The location of the right parenthesis.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal ofLocation As Location, ByVal parameters As IList(Of TypeParameter), ByVal commaLocations As IList(Of Location), ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.TypeParameterCollection, parameters, commaLocations, span)

        _OfLocation = ofLocation
        _RightParenthesisLocation = rightParenthesisLocation
    End Sub
End Class
