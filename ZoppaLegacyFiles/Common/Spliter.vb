Option Strict On
Option Explicit On

Imports System.IO
Imports System.Reflection

''' <summary>文字列分割機能（共通）</summary>
Public MustInherit Class Spliter(Of TItem As {IFileItem})

    ''' <summary>コンストラクタを取得します。</summary>
    Private ReadOnly mConstructor As ConstructorInfo

    ''' <summary>内部ストリームを取得します。</summary>
    Private ReadOnly mInnerStream As TextReader

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="inputStream">入力ストリーム。</param>
    Protected Sub New(inputStream As StreamReader)
        Me.mConstructor = GetType(TItem).GetConstructor(New Type() {GetType(ReadOnlyMemory)})
        Me.mInnerStream = inputStream
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="inputText">入力文字列。</param>
    Protected Sub New(inputText As String)
        Me.mConstructor = GetType(TItem).GetConstructor(New Type() {GetType(ReadOnlyMemory)})
        Me.mInnerStream = New StringReader(inputText)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="inputStream">テキストストリーム。</param>
    Public Sub New(reader As TextReader)
        Me.mConstructor = GetType(TItem).GetConstructor(New Type() {GetType(ReadOnlyMemory)})
        Me.mInnerStream = reader
    End Sub

    ''' <summary>内部より一行を読み込み、分割して返します。</summary>
    ''' <returns>分割した項目の配列。</returns>
    Public Function Split() As List(Of TItem)
        Dim ans = Me.ReadLine()
        Dim res As New List(Of TItem)(ans.Items.Count)
        For Each i In ans.Items
            Dim itm = CType(Me.mConstructor.Invoke(New Object() {i}), TItem)
            res.Add(itm)
        Next
        Return res
    End Function

    ''' <summary>一行読み込み、読み込み結果を取得します。</summary>
    ''' <returns>読み込み結果。</returns>
    Protected Function ReadLine() As ReadResult
        If Me.mInnerStream IsNot Nothing Then
            With Me.ReadLine(Me.mInnerStream)
                Return New ReadResult(.readChars, .splitPoint)
            End With
        Else
            Throw New NullReferenceException("テキストストリームが設定されていません")
        End If
    End Function

    ''' <summary>読み込み結果を指定のデータ型に変換してコンストラクタを実行してインスタンスを返します。</summary>
    ''' <typeparam name="T">インスタンスの型。</typeparam>
    ''' <param name="items">読み込み結果。</param>
    ''' <param name="columTypes">データ型のリスト。</param>
    ''' <param name="constructor">コンストラクタ。</param>
    ''' <returns>インスタンス。</returns>
    Friend Function ReadObject(Of T)(items As ReadOnlyMemory(), columTypes As IDataType(), constructor As ConstructorInfo) As T
        Dim fields As New ArrayList()

        For i As Integer = 0 To Math.Min(columTypes.Length, items.Count) - 1
            Try
                Dim itm = CType(Me.mConstructor.Invoke(New Object() {items(i)}), TItem)
                fields.Add(columTypes(i).ToValue(itm))
            Catch ex As Exception
                Throw New LegacyFilesException($"変換に失敗しました:{i},{items(i)} -> {columTypes(i).DotNetType.Name}")
            End Try
        Next

        Return CType(constructor.Invoke(fields.ToArray()), T)
    End Function

    ''' <summary>一行読み込みを行い、指定のデータ型に変換してコンストラクタを実行してインスタンスを返します。</summary>
    ''' <typeparam name="T">インスタンスの型。</typeparam>
    ''' <param name="columTypes">データ型のリスト。</param>
    ''' <param name="constructor">コンストラクタ。</param>
    ''' <returns>インスタンス。</returns>
    Friend Function InnerReadObject(Of T)(columTypes As IDataType(), constructor As ConstructorInfo) As T
        Dim item = Me.ReadLine()
        If item.HasResult Then
            Return Me.ReadObject(Of T)(item.Items, columTypes, constructor)
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>一行読み込みを行い、分割データを取得します。</summary>
    ''' <returns>分割データ。</returns>
    Public Function ReadFileItems() As TItem()
        Dim item = Me.ReadLine()
        If item.HasResult Then
            Dim res = New TItem(item.Items.Count - 1) {}
            For i As Integer = 0 To item.Items.Count - 1
                res(i) = CType(Me.mConstructor.Invoke(New Object() {item.Items(i)}), TItem)
            Next
            Return res
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>一行読み込みを行い、インスタンスを生成して返します。</summary>
    ''' <typeparam name="T">インスタンスの型。</typeparam>
    ''' <returns>インスタンス。</returns>
    Public Function ReadObject(Of T)() As T
        Dim item = Me.ReadLine()
        If item.HasResult Then
            Return Me.InnerReadObject(Of T)(item.Items).Item1
        End If
        Return Nothing
    End Function

    ''' <summary>一行読み込みを行い、インスタンスを生成して返します（内部用）</summary>
    ''' <typeparam name="T">インスタンスの型。</typeparam>
    ''' <param name="items">読み込んだ情報。</param>
    ''' <returns>インスタンス。</returns>
    Friend Function InnerReadObject(Of T)(items As ReadOnlyMemory()) As (T, ConstructorInfo, IDataType())
        ' 全ての種類のコンストラクタを取得
        Dim constructors = GetType(T).GetConstructors(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)

        For Each con In constructors
            If con.GetParameters().Length = items.Count Then
                Try
                    ' コンストラクタのパラメータから変換するデータ型を取得し、インスタンスを生成
                    Dim columTypes = con.GetParameters().Select(Function(v) DataTypeConverter.Convert(v.ParameterType)).ToArray()
                    Return (Me.ReadObject(Of T)(items, columTypes, con), con, columTypes)
                Catch ex As Exception
                    ' 別のコンストラクタを選択
                End Try
            End If
        Next
        Throw New LegacyFilesException($"変換できるコンストラクタがありません。{String.Join(",", items.Select(Function(s) s.ToString()))}")
    End Function

    ''' <summary>一行読み込みを行い、指定のデータ型に変換してコンストラクタを実行してインスタンスを返します。</summary>
    ''' <typeparam name="T">インスタンスの型。</typeparam>
    ''' <param name="columTypes">データ型のリスト。</param>
    ''' <returns>インスタンス。</returns>
    Public Function ReadObject(Of T)(columTypes As IDataType()) As T
        Dim item = Me.ReadLine()
        If item.HasResult Then
            Return Me.InnerReadObject(Of T)(item.Items, columTypes).Item1
        End If
        Return Nothing
    End Function

    ''' <summary>一行読み込みを行い、指定のデータ型に変換してコンストラクタを実行してインスタンスを返します（内部用）</summary>
    ''' <typeparam name="T">インスタンスの型。</typeparam>
    ''' <param name="columTypes">データ型のリスト。</param>
    ''' <returns>インスタンス。</returns>
    Friend Function InnerReadObject(Of T)(items As ReadOnlyMemory(), columTypes As IDataType()) As (T, ConstructorInfo, IDataType())
        ' 引数の配列を作成
        Dim clmTps = columTypes.Select(Function(v) v.DotNetType).ToArray()

        ' コンストラクタを取得する
        Dim constructor = GetConstructor(Of T)(clmTps)

        ' インスタンスを生成
        Return (Me.ReadObject(Of T)(items, columTypes, constructor), constructor, columTypes)
    End Function

    ''' <summary>コンストラクタの参照を取得します。</summary>
    ''' <typeparam name="TNew">コンストラクタを取得する型。</typeparam>
    ''' <param name="clmTps">コンストラクタの引数。</param>
    ''' <returns>コンストラクタの参照。</returns>
    Private Shared Function GetConstructor(Of TNew)(clmTps() As Type) As ConstructorInfo
        ' コンストラクタの参照を取得
        Dim constructor = GetType(TNew).GetConstructor(clmTps)
        If constructor Is Nothing Then
            constructor = GetType(TNew).GetConstructor(New Type() {GetType(Object())})
        End If

        ' 取得できなければエラー
        If constructor Is Nothing Then
            Dim info = String.Join(",", clmTps.Select(Function(c) c.Name).ToArray())
            Throw New LegacyFilesException($"取得データに一致するコンストラクタがありません:{info}")
        End If
        Return constructor
    End Function

    ''' <summary>一行読み込み、分割を行います。</summary>
    ''' <param name="readStream">読み込みストリーム。</param>
    ''' <returns>読み込んだ文字列と分割位置リスト。</returns>
    Protected MustOverride Function ReadLine(readStream As TextReader) As (readChars As Char(), splitPoint As List(Of Integer))

    ''' <summary>分割結果を保持する。</summary>
    Public Structure ReadResult

        ''' <summary>分割位置リスト。</summary>
        Private ReadOnly mSPoint As List(Of Integer)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="chars">読み込んだ文字配列。</param>
        ''' <param name="spoint">分割位置。</param>
        Public Sub New(chars As Char(), spoint As List(Of Integer))
            Me.Chars = chars
            Me.mSPoint = spoint
        End Sub

        ''' <summary>読み込んだ文字配列を取得します。</summary>
        ''' <returns>文字配列。</returns>
        Public ReadOnly Property Chars() As Char()

        ''' <summary>項目があれば真を返します。</summary>
        ''' <returns>項目があれば真。</returns>
        Public ReadOnly Property HasResult() As Boolean
            Get
                Return (Me.Chars.Length > 0)
            End Get
        End Property

        ''' <summary>項目のリストを返します。</summary>
        ''' <returns>項目リスト。</returns>
        Public ReadOnly Property Items As ReadOnlyMemory()
            Get
                Dim src As New ReadOnlyMemory(Me.Chars)
                Dim split As New List(Of ReadOnlyMemory)(Me.mSPoint.Count - 1)
                For i As Integer = 0 To Me.mSPoint.Count - 2
                    split.Add(src.Slice(Me.mSPoint(i), (Me.mSPoint(i + 1) - 1) - Me.mSPoint(i)))
                Next
                Return split.ToArray()
            End Get
        End Property

    End Structure

End Class
