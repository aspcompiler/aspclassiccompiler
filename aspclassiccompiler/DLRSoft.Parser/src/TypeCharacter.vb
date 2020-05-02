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
''' A character that denotes the type of something.
''' </summary>
<Flags()> _
Public Enum TypeCharacter
    ''' <summary>No type character</summary>
    None = &H0

    ''' <summary>The String symbol '$'.</summary>
    StringSymbol = &H1

    ''' <summary>The Integer symbol '%'.</summary>
    IntegerSymbol = &H2

    ''' <summary>The Long symbol '&amp;'.</summary>
    LongSymbol = &H4

    ''' <summary>The Short character 'S'.</summary>
    ShortChar = &H8

    ''' <summary>The Integer character 'I'.</summary>
    IntegerChar = &H10

    ''' <summary>The Long character 'L'.</summary>
    LongChar = &H20

    ''' <summary>The Single symbol '!'.</summary>
    SingleSymbol = &H40

    ''' <summary>The Double symbol '#'.</summary>
    DoubleSymbol = &H80

    ''' <summary>The Decimal symbol '@'.</summary>
    DecimalSymbol = &H100

    ''' <summary>The Single character 'F'.</summary>
    SingleChar = &H200

    ''' <summary>The Double character 'R'.</summary>
    DoubleChar = &H400

    ''' <summary>The Decimal character 'D'.</summary>
    DecimalChar = &H800

    ''' <summary>The unsigned Short characters 'US'.</summary>
    ''' <remarks>New for Visual Basic 8.0.</remarks>
    UnsignedShortChar = &H1000

    ''' <summary>The unsigned Integer characters 'UI'.</summary>
    ''' <remarks>New for Visual Basic 8.0.</remarks>
    UnsignedIntegerChar = &H2000

    ''' <summary>The unsigned Long characters 'UL'.</summary>
    ''' <remarks>New for Visual Basic 8.0.</remarks>
    UnsignedLongChar = &H4000

    ''' <summary>All type characters.</summary>
    All = &H7FFF
End Enum