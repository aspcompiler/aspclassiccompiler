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
''' The version of the language you want.
''' </summary>
<Flags()> _
Public Enum LanguageVersion
    None = &H0

    ''' <summary>Visual Basic 7.1</summary>
    ''' <remarks>Shipped in Visual Basic 2003</remarks>
    VisualBasic71 = &H1

    ''' <summary>Visual Basic 8.0</summary>
    ''' <remarks>Shipped in Visual Basic 2005</remarks>
    VisualBasic80 = &H2

    All = &H3
End Enum