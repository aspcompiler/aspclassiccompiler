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
''' The type of a unary operator expression.
''' </summary>
Public Enum OperatorType
    None

    ' Unary operators
    UnaryPlus
    Negate
    [Not]

    ' Binary operators
    Plus
    Minus
    Multiply
    Divide
    IntegralDivide
    Concatenate
    ShiftLeft
    ShiftRight
    Power
    Modulus
    [Or]
    [OrElse]
    [And]
    [AndAlso]
    [Xor]
    [Like]
    [Is]
    [IsNot]
    [To]
    Equals
    NotEquals
    LessThan
    LessThanEquals
    GreaterThan
    GreaterThanEquals
End Enum
