Option Strict On
Option Explicit On
Imports System.Text

Namespace Ini

    ''' <summary>キーと値の行を表現します。</summary>
    Friend NotInheritable Class KeyAndValue
        Implements IIniLine

        ''' <summary>行位置を取得します（途中挿入用に小数を指定可）</summary>
        ''' <returns>行位置。</returns>
        Public Property LineNo As Double Implements IIniLine.LineNo

        ''' <summary>書き込み用文字列を取得します。</summary>
        ''' <returns>書き込み文字列。</returns>
        Public ReadOnly Property WriteStr As String Implements IIniLine.WriteStr
            Get
                Return Me.mBaseStr
            End Get
        End Property

        ' エスケープ解除後の文字列
        Private mVUne As String

        ' 値文字列
        Private mVal As String

        ' 定義してあった文字列（元の値）
        Private mBaseStr As String

        ' キーと値の間の位置
        Private ReadOnly mSplitPos As Integer

        ''' <summary>エスケープ解除後の文字列を取得します。</summary>
        Public ReadOnly Property ValueUnEscape As String
            Get
                Return Me.mVUne
            End Get
        End Property

        ''' <summary>値文字列を取得します。</summary>
        Public ReadOnly Property Value As String
            Get
                Return Me.mVal
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="strs">キー／エスケープ解除後の値／値。</param>
        ''' <param name="splitPos">分割位置。</param>
        ''' <param name="row">行位置。</param>
        ''' <param name="str">元の文字列。</param>
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

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="key">キー。</param>
        ''' <param name="val">値。</param>
        ''' <param name="lnno">行位置。</param>
        Public Sub New(key As String, val As String, lnno As Double)
            Me.mVUne = val
            Me.mVal = EscapeValueString(val)
            Dim keystr = EscapeValueString(key)
            Me.LineNo = lnno
            Me.mSplitPos = keystr.Length + 1
            Me.mBaseStr = $"{keystr}={Me.mVal}"
        End Sub

        ''' <summary>値文字列をエスケープします。</summary>
        ''' <param name="input">エスケープする文字列。</param>
        ''' <returns>エスケープされた文字列。</returns>
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

        ''' <summary>値を更新します。</summary>
        ''' <param name="newValue">値文字列。
        ''' </param>
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
