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
''' A collection of case clauses.
''' </summary>
Public NotInheritable Class CaseClauseCollection
    Inherits CommaDelimitedTreeCollection(Of CaseClause)

    ''' <summary>
    ''' Constructs a new collection of case clauses.
    ''' </summary>
    ''' <param name="caseClauses">The case clauses in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the list.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal caseClauses As IList(Of CaseClause), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.CaseClauseCollection, caseClauses, commaLocations, span)

        If caseClauses Is Nothing OrElse caseClauses.Count = 0 Then
            Throw New ArgumentException("CaseClauseCollection cannot be empty.")
        End If
    End Sub
End Class