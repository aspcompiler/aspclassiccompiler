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
''' A parse tree for an Imports statement for a name.
''' </summary>
Public NotInheritable Class NameImport
    Inherits Import

    Private ReadOnly _TypeName As TypeName

    ''' <summary>
    ''' The imported name.
    ''' </summary>
    Public ReadOnly Property TypeName() As TypeName
        Get
            Return _TypeName
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new name import parse tree.
    ''' </summary>
    ''' <param name="typeName">The name to import.</param>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal typeName As TypeName, ByVal span As Span)
        MyBase.New(TreeType.NameImport, span)

        If typeName Is Nothing Then
            Throw New ArgumentNullException("typeName")
        End If

        SetParent(typeName)

        _TypeName = typeName
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, TypeName)
    End Sub
End Class
