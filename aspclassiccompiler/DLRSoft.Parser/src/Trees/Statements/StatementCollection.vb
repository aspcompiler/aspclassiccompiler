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
''' A read-only collection of statements.
''' </summary>
Public NotInheritable Class StatementCollection
    Inherits ColonDelimitedTreeCollection(Of Statement)

    ''' <summary>
    ''' Constructs a new collection of statements.
    ''' </summary>
    ''' <param name="statements">The statements in the collection.</param>
    ''' <param name="colonLocations">The locations of the colons in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal statements As IList(Of Statement), ByVal colonLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.StatementCollection, statements, colonLocations, span)

        If (statements Is Nothing OrElse statements.Count = 0) AndAlso _
           (colonLocations Is Nothing OrElse colonLocations.Count = 0) Then
            Throw New ArgumentException("StatementCollection cannot be empty.")
        End If
    End Sub
End Class