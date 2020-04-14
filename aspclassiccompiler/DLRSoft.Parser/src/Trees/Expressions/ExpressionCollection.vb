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
''' A read-only collection of expressions.
''' </summary>
Public NotInheritable Class ExpressionCollection
    Inherits CommaDelimitedTreeCollection(Of Expression)

    ''' <summary>
    ''' Constructs a new collection of expressions.
    ''' </summary>
    ''' <param name="expressions">The expressions in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal expressions As IList(Of Expression), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.ExpressionCollection, expressions, commaLocations, span)

        If (expressions Is Nothing OrElse expressions.Count = 0) AndAlso _
           (commaLocations Is Nothing OrElse commaLocations.Count = 0) Then
            Throw New ArgumentException("ExpressionCollection cannot be empty.")
        End If
    End Sub
End Class