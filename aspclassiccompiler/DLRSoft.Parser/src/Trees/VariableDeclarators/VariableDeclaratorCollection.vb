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
''' A read-only collection of variable declarators.
''' </summary>
Public NotInheritable Class VariableDeclaratorCollection
    Inherits CommaDelimitedTreeCollection(Of VariableDeclarator)

    ''' <summary>
    ''' Constructs a new collection of variable declarators.
    ''' </summary>
    ''' <param name="variableDeclarators">The variable declarators in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the list.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal variableDeclarators As IList(Of VariableDeclarator), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.VariableDeclaratorCollection, variableDeclarators, commaLocations, span)

        If variableDeclarators Is Nothing OrElse variableDeclarators.Count = 0 Then
            Throw New ArgumentException("VariableDeclaratorCollection cannot be empty.")
        End If
    End Sub
End Class