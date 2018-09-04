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
''' A parse tree for an entire file.
''' </summary>
Public NotInheritable Class File
    Inherits Tree

    Private ReadOnly _Declarations As DeclarationCollection

    ''' <summary>
    ''' The declarations in the file.
    ''' </summary>
    Public ReadOnly Property Declarations() As DeclarationCollection
        Get
            Return _Declarations
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new file parse tree.
    ''' </summary>
    ''' <param name="declarations">The declarations in the file.</param>
    ''' <param name="span">The location of the tree.</param>
    Public Sub New(ByVal declarations As DeclarationCollection, ByVal span As Span)
        MyBase.New(TreeType.File, span)

        SetParent(declarations)

        _Declarations = declarations
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, Declarations)
    End Sub
End Class