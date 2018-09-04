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
''' A read-only collection of imports.
''' </summary>
Public NotInheritable Class ImportCollection
    Inherits CommaDelimitedTreeCollection(Of Import)

    ''' <summary>
    ''' Constructs a collection of imports.
    ''' </summary>
    ''' <param name="importMembers">The imports in the collection.</param>
    ''' <param name="commaLocations">The location of the commas.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal importMembers As IList(Of Import), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.ImportCollection, importMembers, commaLocations, span)

        If importMembers Is Nothing OrElse importMembers.Count = 0 Then
            Throw New ArgumentException("ImportCollection cannot be empty.")
        End If
    End Sub
End Class

