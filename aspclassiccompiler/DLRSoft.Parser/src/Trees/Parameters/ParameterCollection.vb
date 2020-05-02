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
''' A collection of parameters.
''' </summary>
Public NotInheritable Class ParameterCollection
    Inherits CommaDelimitedTreeCollection(Of Parameter)

    Private ReadOnly _RightParenthesisLocation As Location

    ''' <summary>
    ''' The location of the ')'.
    ''' </summary>
    Public ReadOnly Property RightParenthesisLocation() As Location
        Get
            Return _RightParenthesisLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new collection of parameters.
    ''' </summary>
    ''' <param name="parameters">The parameters in the collection</param>
    ''' <param name="commaLocations">The locations of the commas.</param>
    ''' <param name="rightParenthesisLocation">The location of the right parenthesis.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal parameters As IList(Of Parameter), ByVal commaLocations As IList(Of Location), ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.ParameterCollection, parameters, commaLocations, span)

        _RightParenthesisLocation = rightParenthesisLocation
    End Sub
End Class