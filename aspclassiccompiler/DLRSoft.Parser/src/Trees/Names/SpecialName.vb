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
''' A parse tree for a special name (i.e. 'Global').
''' </summary>
Public NotInheritable Class SpecialName
    Inherits Name

    ''' <summary>
    ''' Constructs a new special name parse tree.
    ''' </summary>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal type As TreeType, ByVal span As Span)
        MyBase.New(type, span)

        Debug.Assert(type = TreeType.GlobalNamespaceName OrElse type = TreeType.MeName OrElse type = TreeType.MyBaseName)
    End Sub
End Class