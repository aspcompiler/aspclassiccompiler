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
''' The type of an Option declaration.
''' </summary>
Public Enum OptionType
    SyntaxError
    Explicit
    ExplicitOn
    ExplicitOff
    Strict
    StrictOn
    StrictOff
    CompareBinary
    CompareText
End Enum
