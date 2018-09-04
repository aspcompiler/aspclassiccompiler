'LC Represent a script file
Public Class ScriptBlock
    Inherits Tree

    Private ReadOnly _statements As StatementCollection

    ''' <summary>
    ''' The declarations in the file.
    ''' </summary>
    Public ReadOnly Property Statements() As StatementCollection
        Get
            Return _statements
        End Get
    End Property

    ''' <summary>
    ''' Constructs a new file parse tree.
    ''' </summary>
    ''' <param name="statements">The statements in the file.</param>
    ''' <param name="span">The location of the tree.</param>
    Public Sub New(ByVal statements As StatementCollection, ByVal span As Span)
        MyBase.New(TreeType.ScriptBlock, span)

        SetParent(statements)

        _statements = statements
    End Sub

    Protected Overrides Sub GetChildTrees(ByVal childList As IList(Of Tree))
        AddChild(childList, _statements)
    End Sub
End Class
