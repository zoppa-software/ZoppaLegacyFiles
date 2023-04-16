Option Strict On
Option Explicit On

Namespace Ini

    Friend NotInheritable Class EmptyLine
        Implements IIniLine

        Public Property LineNo As Double Implements IIniLine.LineNo

        Public ReadOnly Property WriteStr As String Implements IIniLine.WriteStr

        Public Sub New(row As Integer, str As String)
            Me.LineNo = row
            Me.WriteStr = str
        End Sub
    End Class

End Namespace