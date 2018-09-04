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
''' A read-only collection of declarations.
''' </summary>
Public NotInheritable Class DeclarationCollection
    Inherits ColonDelimitedTreeCollection(Of Declaration)

    ''' <summary>
    ''' Constructs a new collection of declarations.
    ''' </summary>
    ''' <param name="declarations">The declarations in the collection.</param>
    ''' <param name="colonLocations">The locations of the colons in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal declarations As IList(Of Declaration), ByVal colonLocations As IList(Of Location), ByVal span As Span)
        MyBase.New(TreeType.DeclarationCollection, declarations, colonLocations, span)

        ' A declaration collection may need to hold just a colon.
        If (declarations Is Nothing OrElse declarations.Count = 0) AndAlso _
           (colonLocations Is Nothing OrElse colonLocations.Count = 0) Then
            Throw New ArgumentException("DeclarationCollection cannot be empty.")
        End If
    End Sub
End Class