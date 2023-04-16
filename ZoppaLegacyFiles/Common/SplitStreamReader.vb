Option Strict On
Option Explicit On

Imports System.IO
Imports System.Reflection

''' <summary>読み込みストリーム（共通）</summary>
Public MustInherit Class SplitStreamReader(Of TSpliter As {Spliter(Of TItem)}, TItem As {IFileItem})
    Inherits IO.StreamReader
    Implements IEnumerable(Of Pointer)

    ''' <summary>コンストラクタを取得します。</summary>
    Private mConstructor As ConstructorInfo

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    Public Sub New(stream As IO.Stream)
        MyBase.New(stream)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    Public Sub New(path As String)
        MyBase.New(path)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(stream As IO.Stream, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(stream, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding)
        MyBase.New(stream, encoding)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(path As String, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(path, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    Public Sub New(path As String, encoding As Text.Encoding)
        MyBase.New(path, encoding)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean)
        MyBase.New(path, encoding, detectEncodingFromByteOrderMarks)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
        MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="path">入力ファイルパス。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    Public Sub New(path As String, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer)
        MyBase.New(path, encoding, detectEncodingFromByteOrderMarks, bufferSize)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="stream">元となるストリーム。</param>
    ''' <param name="encoding">テキストエンコード。</param>
    ''' <param name="detectEncodingFromByteOrderMarks">バイトオーダーマーク。</param>
    ''' <param name="bufferSize">バッファサイズ。</param>
    ''' <param name="leaveOpen">ストリームを開いたままにするならば真。</param>
    Public Sub New(stream As IO.Stream, encoding As Text.Encoding, detectEncodingFromByteOrderMarks As Boolean, bufferSize As Integer, leaveOpen As Boolean)
        MyBase.New(stream, encoding, detectEncodingFromByteOrderMarks, bufferSize, leaveOpen)
    End Sub

    ''' <summary>列挙子を取得する。</summary>
    ''' <returns>列挙子。</returns>
    Public Iterator Function GetEnumerator() As IEnumerator(Of Pointer) Implements IEnumerable(Of Pointer).GetEnumerator
        ' 分割機能のコンストラクタを取得
        If Me.mConstructor Is Nothing Then
            Me.mConstructor = GetType(TSpliter).GetConstructor(New Type() {GetType(StreamReader)})
        End If

        ' 分割機能を生成
        Dim spliter = CType(Me.mConstructor.Invoke(New Object() {Me}), TSpliter)

        ' 項目を分割して返す
        Dim ans = spliter.ReadFileItems()
        Dim index As Integer = 0
        Do While ans IsNot Nothing
            Yield New Pointer(index, ans)
            index += 1
            ans = spliter.ReadFileItems()
        Loop
    End Function

    ''' <summary>列挙処理の前処理を行う。</summary>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <returns>列挙子と分割機能。</returns>
    Private Function PreProcessing(topSkip As Integer) As (iter As IEnumerator(Of Pointer), spliter As TSpliter)
        ' 列挙子を取得
        Dim iter = Me.GetEnumerator()

        ' スキップ数を読み捨て
        Do While iter.MoveNext()
            Dim cur = iter.Current
            If cur.Row >= topSkip Then
                Exit Do
            End If
        Loop

        ' 分割機能を取得
        Dim spliter = CType(Me.mConstructor.Invoke(New Object() {Me}), TSpliter)

        Return (iter, spliter)
    End Function

    ''' <summary>一行の情報を指定の型に変換して取得します。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <param name="rowLimit">読み込む行数。負数の場合は制限なし。</param>
    ''' <returns>変換後の型の列挙子。</returns>
    Public Iterator Function [Select](Of T)(Optional topSkip As Integer = 0, Optional rowLimit As Integer = -1) As IEnumerable(Of T)
        ' 列挙子と分割機能を取得
        Dim pair = Me.PreProcessing(topSkip)

        Dim constructor As ConstructorInfo = Nothing
        Dim dataTypes As IDataType() = Nothing

        ' 1. カレントの情報を取得
        ' 2. 制限なしか指定行数内ならば処理する
        ' 3. カレント情報の分割項目を取得
        ' 4. 変換後の型のコンストラクタを取得していない場合、コンストラクタを取得して取得
        ' 5. 取得している場合、コンストラクタから取得
        Do
            Dim cur = pair.iter.Current                                     ' 1

            If rowLimit < 0 OrElse cur.Row < topSkip + rowLimit Then        ' 2
                Dim items = cur.Items.Select(Function(r) r.Raw).ToArray()   ' 3

                If constructor Is Nothing Then
                    Dim dat = pair.spliter.InnerReadObject(Of T)(items)     ' 4
                    constructor = dat.Item2
                    dataTypes = dat.Item3
                    Yield dat.Item1
                Else
                    Yield pair.spliter.ReadObject(Of T)(items, dataTypes, constructor)  ' 5
                End If
            Else
                Exit Do
            End If
        Loop While pair.iter.MoveNext()
    End Function

    ''' <summary>一行の情報を指定の型に変換して取得します。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="convertfunc">変換処理。</param>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <param name="rowLimit">読み込む行数。負数の場合は制限なし。</param>
    ''' <returns>変換後の型の列挙子。</returns>
    Public Iterator Function [Select](Of T)(convertfunc As Func(Of Integer, TItem(), T),
                                            Optional topSkip As Integer = 0, Optional rowLimit As Integer = -1) As IEnumerable(Of T)
        ' 列挙子と分割機能を取得
        Dim pair = Me.PreProcessing(topSkip)

        Dim constructor As ConstructorInfo = Nothing
        Dim dataTypes As IDataType() = Nothing

        ' 1. カレントの情報を取得
        ' 2. 制限なしか指定行数内ならば処理する
        ' 3. 変換メソッドを使用して取得
        Do
            Dim cur = pair.iter.Current                                 ' 1
            If rowLimit < 0 OrElse cur.Row < topSkip + rowLimit Then    ' 2
                Yield convertfunc(cur.Row, cur.Items)                   ' 3
            Else
                Exit Do
            End If
        Loop While pair.iter.MoveNext()
    End Function

    ''' <summary>指定の型に変換して取得します。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="dataTypes">各列の型リスト。</param>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <param name="rowLimit">読み込む行数。負数の場合は制限なし。</param>
    ''' <returns>変換後の型の列挙子。</returns>
    Public Iterator Function [Select](Of T)(dataTypes As IEnumerable(Of IDataType),
                                            Optional topSkip As Integer = 0, Optional rowLimit As Integer = -1) As IEnumerable(Of T)
        ' 列挙子と分割機能を取得
        Dim pair = Me.PreProcessing(topSkip)

        Dim constructor As ConstructorInfo = Nothing
        Dim dataTypeArray = dataTypes.ToArray()

        ' 1. カレントの情報を取得
        ' 2. 制限なしか指定行数内ならば処理する
        ' 3. カレント情報の分割項目を取得
        ' 4. 変換後の型のコンストラクタを取得していない場合、コンストラクタを取得して取得
        ' 5. 取得している場合、コンストラクタから取得
        Do
            Dim cur = pair.iter.Current                                                 ' 1

            If rowLimit < 0 OrElse cur.Row < topSkip + rowLimit Then                    ' 2
                Dim items = cur.Items.Select(Function(r) r.Raw).ToArray()               ' 3

                If constructor Is Nothing Then
                    Dim dat = pair.spliter.InnerReadObject(Of T)(items, dataTypeArray)  ' 4
                    constructor = dat.Item2
                    Yield dat.Item1
                Else
                    Yield pair.spliter.ReadObject(Of T)(items, dataTypeArray, constructor)  ' 5
                End If
            Else
                Exit Do
            End If
        Loop While pair.iter.MoveNext()
    End Function

    ''' <summary>条件を満たす行を指定の型に変換して取得します。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="condition">条件判定処理。</param>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <param name="rowLimit">読み込む行数。負数の場合は制限なし。</param>
    ''' <returns>変換後の型の列挙子。</returns>
    Public Iterator Function Where(Of T)(condition As Func(Of Integer, TItem(), Boolean),
                                         Optional topSkip As Integer = 0, Optional rowLimit As Integer = -1) As IEnumerable(Of T)
        ' 列挙子と分割機能を取得
        Dim pair = Me.PreProcessing(topSkip)

        Dim constructor As ConstructorInfo = Nothing
        Dim dataTypes As IDataType() = Nothing

        ' 1. カレントの情報を取得
        ' 2. 制限なしか指定行数内ならば処理する
        ' 3. カレント情報の分割項目を取得
        ' 4. 変換後の型のコンストラクタを取得していない場合、コンストラクタを取得して取得
        ' 5. 取得している場合、コンストラクタから取得
        Do
            Dim cur = pair.iter.Current                                         ' 1

            If rowLimit < 0 OrElse cur.Row < topSkip + rowLimit Then            ' 2
                If condition(cur.Row, cur.Items) Then
                    Dim items = cur.Items.Select(Function(r) r.Raw).ToArray()   ' 3

                    If constructor Is Nothing Then
                        Dim dat = pair.spliter.InnerReadObject(Of T)(items)     ' 4
                        constructor = dat.Item2
                        dataTypes = dat.Item3
                        Yield dat.Item1
                    Else
                        Yield pair.spliter.ReadObject(Of T)(items, dataTypes, constructor)  ' 5
                    End If
                End If
            Else
                Exit Do
            End If
        Loop While pair.iter.MoveNext()
    End Function

    ''' <summary>条件を満たす行を指定の型に変換して取得します。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="condition">条件判定処理。</param>
    ''' <param name="convertfunc">変換処理。</param>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <param name="rowLimit">読み込む行数。負数の場合は制限なし。</param>
    ''' <returns>変換後の型の列挙子。</returns>
    Public Iterator Function Where(Of T)(condition As Func(Of Integer, TItem(), Boolean),
                                         convertfunc As Func(Of Integer, TItem(), T),
                                         Optional topSkip As Integer = 0, Optional rowLimit As Integer = -1) As IEnumerable(Of T)
        ' 列挙子と分割機能を取得
        Dim pair = Me.PreProcessing(topSkip)

        Dim constructor As ConstructorInfo = Nothing
        Dim dataTypes As IDataType() = Nothing

        ' 1. カレントの情報を取得
        ' 2. 制限なしか指定行数内ならば処理する
        ' 3. 変換メソッドを使用して取得
        Do
            Dim cur = pair.iter.Current                                 ' 1
            If rowLimit < 0 OrElse cur.Row < topSkip + rowLimit Then    ' 2
                If condition(cur.Row, cur.Items) Then
                    Yield convertfunc(cur.Row, cur.Items)               ' 3
                End If
            Else
                Exit Do
            End If
        Loop While pair.iter.MoveNext()
    End Function

    ''' <summary>条件を満たす行を指定の型に変換して取得します。</summary>
    ''' <typeparam name="T">変換後の型。</typeparam>
    ''' <param name="condition">条件判定処理。</param>
    ''' <param name="dataTypes">各列の型リスト。</param>
    ''' <param name="topSkip">最初からの読み捨て行数。</param>
    ''' <param name="rowLimit">読み込む行数。負数の場合は制限なし。</param>
    ''' <returns>変換後の型の列挙子。</returns>
    Public Iterator Function Where(Of T)(condition As Func(Of Integer, TItem(), Boolean),
                                         dataTypes As IEnumerable(Of IDataType),
                                         Optional topSkip As Integer = 0, Optional rowLimit As Integer = -1) As IEnumerable(Of T)
        ' 列挙子と分割機能を取得
        Dim pair = Me.PreProcessing(topSkip)

        Dim constructor As ConstructorInfo = Nothing
        Dim dataTypeArray = dataTypes.ToArray()

        ' 1. カレントの情報を取得
        ' 2. 制限なしか指定行数内ならば処理する
        ' 3. カレント情報の分割項目を取得
        ' 4. 変換後の型のコンストラクタを取得していない場合、コンストラクタを取得して取得
        ' 5. 取得している場合、コンストラクタから取得
        Do
            Dim cur = pair.iter.Current                                                     ' 1

            If rowLimit < 0 OrElse cur.Row < topSkip + rowLimit Then                        ' 2
                If condition(cur.Row, cur.Items) Then
                    Dim items = cur.Items.Select(Function(r) r.Raw).ToArray()               ' 3

                    If constructor Is Nothing Then
                        Dim dat = pair.spliter.InnerReadObject(Of T)(items, dataTypeArray)  ' 4
                        constructor = dat.Item2
                        Yield dat.Item1
                    Else
                        Yield pair.spliter.ReadObject(Of T)(items, dataTypeArray, constructor)  ' 5
                    End If
                End If
            Else
                Exit Do
            End If
        Loop While pair.iter.MoveNext()
    End Function

    ''' <summary>列挙子を取得する。</summary>
    ''' <returns>列挙子。</returns>
    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    ''' <summary>項目のポインタ。</summary>
    Public Structure Pointer

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="row">行位置。</param>
        ''' <param name="items">分割項目。</param>
        Public Sub New(row As Integer, items As TItem())
            Me.Row = row
            Me.Items = items
        End Sub

        ''' <summary>行位置を取得する。</summary>
        ''' <returns>行位置。</returns>
        Public ReadOnly Property Row() As Integer

        ''' <summary>分割項目の配列を取得する。</summary>
        ''' <returns>分割項目配列。</returns>
        Public ReadOnly Property Items() As TItem()

        ''' <summary>文字表現を取得する。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return $"row index:{Me.Row} colum count:{Me.Items.Length}"
        End Function

    End Structure

End Class
