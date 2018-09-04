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
''' A collection of type arguments.
''' </summary>
Public NotInheritable Class TypeArgumentCollection
    Inherits CommaDelimitedTreeCollection(Of TypeName)

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
    ''' Constructs a new collection of type arguments.
    ''' </summary>
    ''' <param name="ofLocation">The location of the 'Of'.</param>
    ''' <param name="arguments">The type arguments in the collection</param>
    ''' <param name="commaLocations">The locations of the commas.</param>
    ''' <param name="rightParenthesisLocation">The location of the right parenthesis.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal ofLocation As Location, ByVal arguments As IList(Of TypeName), ByVal commaLocations As IList(Of Location), ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.TypeArgumentCollection, arguments, commaLocations, span)

        _OfLocation = ofLocation
        _RightParenthesisLocation = rightParenthesisLocation
    End Sub
End Class
