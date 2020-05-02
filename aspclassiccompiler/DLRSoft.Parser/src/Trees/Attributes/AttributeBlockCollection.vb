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
''' A read-only collection of attributes.
''' </summary>
Public NotInheritable Class AttributeBlockCollection
    Inherits TreeCollection(Of AttributeCollection)

    ''' <summary>
    ''' Constructs a new collection of attribute blocks.
    ''' </summary>
    ''' <param name="attributeBlocks">The attribute blockss in the collection.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal attributeBlocks As IList(Of AttributeCollection), ByVal span As Span)
        MyBase.New(TreeType.AttributeBlockCollection, attributeBlocks, span)

        If attributeBlocks Is Nothing OrElse attributeBlocks.Count = 0 Then
            Throw New ArgumentException("AttributeBlocksCollection cannot be empty.")
        End If
    End Sub
End Class
