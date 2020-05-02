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
''' The type of a parse tree modifier.
''' </summary>
<Flags()> Public Enum ModifierTypes
    None = &H0
    [Public] = &H1
    [Private] = &H2
    [Protected] = &H4
    [Friend] = &H8
    AccessModifiers = [Public] Or [Private] Or [Protected] Or [Friend]
    [Static] = &H10
    [Shared] = &H20
    [Shadows] = &H40
    [Overloads] = &H80
    [MustInherit] = &H100
    [NotInheritable] = &H200
    [Overrides] = &H400
    [NotOverridable] = &H800
    [Overridable] = &H1000
    [MustOverride] = &H2000
    [ReadOnly] = &H4000
    [WriteOnly] = &H8000
    [Dim] = &H10000
    [Const] = &H20000
    [Default] = &H40000
    [WithEvents] = &H80000
    [ByVal] = &H100000
    [ByRef] = &H200000
    [Optional] = &H400000
    [ParamArray] = &H800000
    [Partial] = &H1000000
    [Widening] = &H2000000
    [Narrowing] = &H4000000
End Enum