Option Strict On
Option Explicit On
Imports System.Text

Namespace Ini

    Friend NotInheritable Class KeyAndValue
        Implements IIniLine

        Public Property LineNo As Double Implements IIniLine.LineNo

        Public ReadOnly Property WriteStr As String Implements IIniLine.WriteStr
            Get
                Return Me.mBaseStr
            End Get
        End Property

        Private mVUne As String

        Private mVal As String

        Private mBaseStr As String

        Private mSplitPos As Integer

        Public ReadOnly Property ValueUnEscape As String
            Get
                Return Me.mVUne
            End Get
        End Property

        Public ReadOnly Property Value As String
            Get
                Return Me.mVal
            End Get
        End Property

        Public Sub New(strs As String(), splitPos As Integer, row As Integer, str As String)
            Me.mVUne = strs(2)
            Me.mVal = strs(3)
            Me.LineNo = row
            Me.mBaseStr = str
            Dim pad As Integer = 0
            For Each c In str.ToCharArray()
                If Char.IsWhiteSpace(c) Then
                    pad += 1
                Else
                    Exit For
                End If
            Next
            Me.mSplitPos = splitPos + pad
        End Sub

        Private Shared Function EscapeValueString(input As String) As String
            Dim quote = (input.Trim() <> input)
            Dim buf As New StringBuilder()
            If quote Then
                buf.Append(""""c)
            End If
            For Each c In input.ToCharArray()
                Select Case c
                    Case ";"c, "#"c, "="c, ":"c, "\"c
                        buf.Append("\"c)
                        buf.Append(c)
                    Case """"c
                        If quote Then
                            buf.Append(c)
                        End If
                        buf.Append(c)
                    Case CChar(vbNullChar)
                        buf.Append("\0")
                    Case CChar(vbTab)
                        buf.Append("\t")
                    Case CChar(vbCr)
                        buf.Append("\r")
                    Case CChar(vbLf)
                        buf.Append("\n")
                    Case Else
                        buf.Append(c)
                End Select
            Next
            If quote Then
                buf.Append(""""c)
            End If
            Return buf.ToString()
        End Function

        Public Sub New(key As String, val As String, lnno As Double)
            Me.mVUne = val
            Me.mVal = EscapeValueString(val)
            Dim keystr = EscapeValueString(key)
            Me.LineNo = lnno
            Me.mSplitPos = keystr.Length + 1
            Me.mBaseStr = $"{keystr}={Me.mVal}"
        End Sub

        Public Sub UpdateValue(newValue As String)
            Dim keyStr = Me.mBaseStr.Substring(0, Me.mSplitPos)
            Dim comPos = Me.mSplitPos + Me.mVal.Length
            Dim comStr = Me.mBaseStr.Substring(comPos, Me.mBaseStr.Length - comPos)

            Me.mVUne = newValue
            Me.mVal = EscapeValueString(newValue)
            Me.mBaseStr = keyStr & Me.mVal & comStr
        End Sub

    End Class

End Namespace
