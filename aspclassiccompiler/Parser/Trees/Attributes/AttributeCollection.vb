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
''' A read-only collection of attributes.
''' </summary>
Public NotInheritable Class AttributeCollection
    Inherits CommaDelimitedTreeCollection(Of Attribute)

    Private ReadOnly _RightBracketLocation As Location

    ''' <summary>
    ''' The location of the '}'.
    ''' </summary>
    Public ReadOnly Property RightBracketLocation() As Location
        Get
            Return _RightBracketLocation
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new collection of attributes.
    ''' </summary>
    ''' <param name="attributes">The attributes in the collection.</param>
    ''' <param name="commaLocations">The location of the commas in the list.</param>
    ''' <param name="rightBracketLocation">The location of the right bracket.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal attributes As IList(Of Attribute), ByVal commaLocations As IList(Of Location), ByVal rightBracketLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.AttributeCollection, attributes, commaLocations, span)

        If attributes Is Nothing OrElse attributes.Count = 0 Then
            Throw New ArgumentException("AttributeCollection cannot be empty.")
        End If

        _RightBracketLocation = rightBracketLocation
    End Sub
End Class
