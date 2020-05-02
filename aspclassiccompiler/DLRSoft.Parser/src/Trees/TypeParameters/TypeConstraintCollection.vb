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
''' A collection of type constraints.
''' </summary>
Public NotInheritable Class TypeConstraintCollection
    Inherits CommaDelimitedTreeCollection(Of TypeName)

    Private ReadOnly _RightBracketLocation As Location

    ''' <summary>
    ''' The location of the '}', if any.
    ''' </summary>
    Public ReadOnly Property RightBracketLocation() As Location
        Get
            Return _RightBracketLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new collection of type constraints.
    ''' </summary>
    ''' <param name="constraints">The type constraints in the collection</param>
    ''' <param name="commaLocations">The locations of the commas.</param>
    ''' <param name="rightBracketLocation">The location of the right bracket, if any.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal constraints As IList(Of TypeName), ByVal commaLocations As IList(Of Location), ByVal rightBracketLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.TypeConstraintCollection, constraints, commaLocations, span)

        _RightBracketLocation = rightBracketLocation
    End Sub
End Class
