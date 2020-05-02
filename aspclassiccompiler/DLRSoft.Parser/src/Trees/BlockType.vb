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
''' The type a block declaration.
''' </summary>
Public Enum BlockType
    None

    [Do]
    [For]

    [While]
    [Select]
    [If]
    [Try]
    [SyncLock]
    [Using]
    [With]

    [Sub]
    [Function]
    [Operator]

    [Event]
    [AddHandler]
    [RemoveHandler]
    [RaiseEvent]

    [Get]
    [Set]
    [Property]

    [Class]
    [Structure]
    [Module]
    [Interface]
    [Enum]
    [Namespace]
End Enum