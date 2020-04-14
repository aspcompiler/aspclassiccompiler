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
''' A parse tree for a type name.
''' </summary>
Public MustInherit Class TypeName
    Inherits Tree

    Protected Sub New(ByVal type As TreeType, ByVal span As Span)
        MyBase.New(type, span)

        Debug.Assert(type >= TreeType.IntrinsicType AndAlso type <= TreeType.ArrayType)
    End Sub
End Class