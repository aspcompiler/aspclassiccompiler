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
''' A read-only collection of variable names.
''' </summary>
Public NotInheritable Class VariableNameCollection
    Inherits CommaDelimitedTreeCollection(Of VariableName)
    ''' <summary>
    ''' Constructs a new variable name collection.
    ''' </summary>
    ''' <param name="variableNames">The variable names in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal variableNames As IList(Of VariableName), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.VariableNameCollection, variableNames, commaLocations, span)

        If variableNames Is Nothing OrElse variableNames.Count = 0 Then
            Throw New ArgumentException("VariableNameCollection cannot be empty.")
        End If
    End Sub
End Class