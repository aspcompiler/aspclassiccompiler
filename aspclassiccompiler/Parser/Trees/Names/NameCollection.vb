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
''' A read-only collection of names.
''' </summary>
Public NotInheritable Class NameCollection
    Inherits CommaDelimitedTreeCollection(Of Name)
    ''' <summary>
    ''' Constructs a new name collection.
    ''' </summary>
    ''' <param name="names">The names in the collection.</param>
    ''' <param name="commaLocations">The locations of the commas in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal names As IList(Of Name), ByVal commaLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.NameCollection, names, CommaLocations, Span)

        If names Is Nothing OrElse names.Count = 0 Then
            Throw New ArgumentException("NameCollection cannot be empty.")
        End If
    End Sub
End Class