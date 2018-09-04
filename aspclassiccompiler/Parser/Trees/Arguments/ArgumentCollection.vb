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
''' A read-only collection of arguments.
''' </summary>
Public NotInheritable Class ArgumentCollection
    Inherits CommaDelimitedTreeCollection(Of Argument)

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
    ''' Constructs a new argument collection.
    ''' </summary>
    ''' <param name="arguments">The arguments in the collection.</param>
    ''' <param name="commaLocations">The location of the commas in the collection.</param>
    ''' <param name="rightParenthesisLocation">The location of the ')'.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal arguments As IList(Of Argument), ByVal commaLocations As IList(Of Location), ByVal rightParenthesisLocation As Location, ByVal span As Span)
        MyBase.New(TreeType.ArgumentCollection, arguments, commaLocations, span)

        _RightParenthesisLocation = rightParenthesisLocation
    End Sub
End Class