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
''' A parse tree for the global namespace (i.e. 'Global').
''' </summary>
Public NotInheritable Class GlobalNamespaceName
    Inherits Name

    ''' <summary>
    ''' Constructs a new global namespace name parse tree.
    ''' </summary>
    ''' <param name="span">The location of the parse tree.</param>
    Public Sub New(ByVal span As Span)
        MyBase.New(TreeType.GlobalNamespaceName, span)
    End Sub
End Class