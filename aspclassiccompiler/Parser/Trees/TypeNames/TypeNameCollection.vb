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
''' A read-only collection of type names.
''' </summary>
Public NotInheritable Class TypeNameCollection
    Inherits CommaDelimitedTreeCollection(Of TypeName)

    ''' <summary>
    ''' Constructs a new type name collection.
    ''' </summary>
    ''' <param name="typeMembers">The type names in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal typeMembers As IList(Of TypeName), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.TypeNameCollection, typeMembers, commaLocations, span)

        If typeMembers Is Nothing OrElse typeMembers.Count = 0 Then
            Throw New ArgumentException("TypeNameCollection cannot be empty.")
        End If
    End Sub
End Class