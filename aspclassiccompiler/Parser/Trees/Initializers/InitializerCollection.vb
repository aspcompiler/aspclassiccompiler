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
''' A read-only collection of initializers.
''' </summary>
Public NotInheritable Class InitializerCollection
    Inherits CommaDelimitedTreeCollection(Of Initializer)
    Private ReadOnly _RightCurlyBraceLocation As Location

    ''' <summary>
    ''' The location of the '}'.
    ''' </summary>
    Public ReadOnly Property RightCurlyBraceLocation() As Location
        Get
            Return _RightCurlyBraceLocation
        End Get
    End Property
    ''' <summary>
    ''' Constructs a new initializer collection.
    ''' </summary>
    ''' <param name="initializers">The initializers in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the collection.</param>
    ''' <param name="rightCurlyBraceLocation">The location of the '}'.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal initializers As IList(Of Initializer), ByVal commaLocations As IList(Of Location), ByVal rightCurlyBraceLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.InitializerCollection, initializers, commaLocations, span)

        _RightCurlyBraceLocation = rightCurlyBraceLocation
    End Sub
End Class