

Imports System.Timers
Imports System.Threading
Imports System.Runtime.InteropServices

#If 1 Then
Imports canlibCLSNET
#Else
'Imports Kvaser.CanLib.Canlib
Imports Kvaser.CanLib
#End If

''' <summary>
''' ウィンドウメッセージをトリガにするためのラッパクラス
''' </summary>
''' <remarks>NativeWindowの基本的な実装になる</remarks>
Public Class MessageCapture
    Inherits NativeWindow

    '''<summary>ウィンドウ非表示のためのスタイル指定</summary>
    Private Const WS_BORDER As Int32 = &H800000

    ''' <summary>
    ''' メッセージプロシージャのデリゲート
    ''' </summary>
    Public Delegate Function MsgProc(ByRef m As System.Windows.Forms.Message) As Boolean
    Private fProc As MsgProc

    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <param name="proc">メッセージ受信時の処理デリゲート</param>
    Public Sub New(ByVal proc As MsgProc)
        MyBase.New()
        'デリゲートを保存
        fProc = proc
        'ハンドル登録用のWindow設定
        Dim cp As New CreateParams
        cp.X = 0
        cp.Y = 0
        cp.Height = 0
        cp.Width = 0
        cp.Style = WS_BORDER
        'ウィンドウハンドルの作成・登録
        Me.CreateHandle(cp)

    End Sub

    ''' <summary>
    ''' デストラクタ
    ''' </summary>
    Protected Overrides Sub Finalize()
        'ウィンドウハンドル他の解放
        Me.DestroyHandle()
    End Sub

    ''' <summary>
    ''' ウィンドウプロシージャ
    ''' </summary>
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        '戻り値によってはBaseクラスへ処理を継続しない
        If Not fProc(m) Then
            Exit Sub
        End If
        MyBase.WndProc(m)

    End Sub

End Class


''' <summary>
''' CAN通信クラス
''' Kvaser Leaf Light(USB)用
''' </summary>
''' <remarks></remarks>
Public Class CAN_Kvaser


    ''' <summary>
    ''' CAN通信ハンドルナンバー
    ''' </summary>
    ''' <remarks></remarks>
    Private canHandle As Integer

    ''' <summary>
    ''' １ビットのクオンタム設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Const bitQuantams As Integer = 16

    ''' <summary>
    ''' タイムセグメント2
    ''' </summary>
    ''' <remarks></remarks>
    Private Const tseg2 As Integer = 4

    ''' <summary>
    ''' タイムセグメント1
    ''' </summary>
    ''' <remarks></remarks>
    Private Const tseg1 As Integer = (bitQuantams - tseg2 - 1)

    ''' <summary>
    ''' シンクロナイゼーションジャンプ
    ''' </summary>
    ''' <remarks></remarks>
    Private Const sjw As Integer = 3

    ''' <summary>
    ''' CAN通信速度
    ''' </summary>
    ''' <remarks></remarks>
    Private canBaud As Integer = Canlib.canBITRATE_500K

    ''' <summary>
    ''' kvaserからの受信割り込み用
    ''' </summary>
    ''' <remarks></remarks>
    Private fMsgCapture As MessageCapture

    ''' <summary>
    ''' データ受信イベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event OnComm(ByVal sender As Object, ByVal e As myEventArgs)

    ''' <summary>
    ''' 送受信成功フラグ
    ''' </summary>
    ''' <remarks></remarks>
    Private TxRx_Rcv As Boolean

    ''' <summary>
    ''' 受信データ
    ''' </summary>
    ''' <remarks>読み出す場合はGetRxDatメソッドを使う</remarks>
    Private rcv_can_data As CANDFRAME

    ''' <summary>
    ''' 特別な受信ID
    ''' </summary>
    ''' <remarks></remarks>
    Public specialID As New List(Of Integer)


    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        rcv_can_data = New CANDFRAME

    End Sub


    ''' <summary>
    ''' CANデバイスの探索
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeviceSearch() As String()

        Dim ports As UInteger
        Dim channels As New List(Of String)

        Canlib.canUnloadLibrary()

        'Kvaserポートを探す
        Try
            Canlib.canInitializeLibrary()
            If Canlib.canGetNumberOfChannels(ports) = Canlib.canStatus.canOK And ports > 0 Then
                For i = 0 To ports - 1
                    Dim objBuf As New Object
                    Canlib.canGetChannelData(i, Canlib.canCHANNELDATA_CARD_TYPE, objBuf)
                    Dim portstr As String = ""
                    If CType(objBuf, Integer) <> Canlib.canHWTYPE_VIRTUAL Then
                        portstr = "physical"
                    Else
                        portstr = "virtual"
                    End If

                    channels.Add(i & ":" & portstr)
                    'channels.Add(i & ":" & IIf(objBuf <> Canlib.canHWTYPE_VIRTUAL, "physical", "virtual"))
                Next
            End If
        Catch ex As Exception

        End Try

        Return IIf(channels.Count > 0, channels.ToArray, Nothing)

    End Function


    ''' <summary>
    ''' ポートオープンとCAN通信速度ビット設定
    ''' </summary>
    ''' <param name="port"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function open(ByVal port As String) As Boolean

        If port.Contains("physical") Or port.Contains("virtual") Then

            Canlib.canInitializeLibrary()

            Canlib.canClose(canHandle)
            canHandle = Canlib.canOpenChannel(Split(port, ":")(0), Canlib.canOPEN_ACCEPT_VIRTUAL)

            If canHandle >= 0 Then
                If Canlib.canSetBusParams(canHandle, Canlib.canBITRATE_500K, tseg1, tseg2, sjw, 1, 0) = Canlib.canStatus.canOK Then
                    If Canlib.canBusOn(canHandle) = Canlib.canStatus.canOK Then

                        '受信メッセージキャプチャの設定(空の非表示ウィンドウに受信時の処理関数を割当)
                        fMsgCapture = New MessageCapture(AddressOf WndMsgProc)

                        '受信イベントの設定
                        If Canlib.canSetNotify(canHandle, fMsgCapture.Handle, Canlib.canNOTIFY_RX) = Canlib.canStatus.canOK Then
                            Return True
                        End If

                    End If
                End If
            End If

            Canlib.canClose(canHandle)

        Else

        End If

        Return False

    End Function


    ''' <summary>
    ''' ポートのクローズ
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub close()

        'KVASERを閉じる
        Try
            Canlib.canBusOff(canHandle)
            Canlib.canClose(canHandle)
        Catch ex As Exception

        End Try

    End Sub


    ''' <summary>
    ''' CANデータの送信
    ''' </summary>
    ''' <param name="cand"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Tx(ByVal cand As CANDFRAME) As Boolean
        Return kvaser_CAN_TX(cand)
    End Function


    ''' <summary>
    ''' KVASRでの送受信関数
    ''' </summary>
    ''' <param name="txd"></param>
    ''' <param name="rxd"></param>
    ''' <param name="timeout"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TxRx_kvaser(ByVal txd As CANDFRAME, ByRef rxd As CANDFRAME, ByVal timeout As Long) As Boolean

        If timeout < 5 Then
            timeout = 5
        End If

        Me.TxRx_Rcv = False
        If kvaser_CAN_TX(txd, timeout) Then

            Dim ktimer As New Timers.Timer
            ktimer.AutoReset = False
            ktimer.Interval = timeout
            ktimer.Enabled = True
            Do
                Rx()
            Loop While (Me.TxRx_Rcv = False) And (ktimer.Enabled = True)
            ktimer.Enabled = False
            ktimer.Dispose()

            If Me.TxRx_Rcv Then
                rxd = GetRxDat()
                Return True
            End If
        End If

        Return False

    End Function


    ''' <summary>
    ''' 応答を期待する送信シークエンス
    ''' </summary>
    ''' <param name="txd">送信フレーム</param>  
    ''' <param name="rxd">応答フレーム</param>
    ''' <param name="timeout">タイムアウト時間[ms]</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TxRx(ByVal txd As CANDFRAME, ByVal timeout As UShort, ByRef rxd As CANDFRAME) As Boolean

        If TxRx_kvaser(txd, rxd, timeout) Then
            If Me.specialID.Contains(rxd.id) Then
                Return True
            End If
        End If
        Return False

    End Function


    ''' <summary>
    ''' 文字列中に含まれる指定文字の個数を返す
    ''' </summary>
    ''' <param name="str"></param>
    ''' <param name="cc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function countChar(ByVal str As String, ByVal cc As Char) As Integer

        Return str.Length - str.Replace(cc, "").Length

    End Function


    ''' <summary>
    ''' 最新のデバッグ受信データを返す
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRxDat() As CANDFRAME

        Return rcv_can_data

    End Function


    Private lock_kvaser_CAN_TX As New Object


    ''' <summary>
    ''' KvaserでCANデータ送信
    ''' </summary>
    ''' <param name="cand"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function kvaser_CAN_TX(ByVal cand As CANDFRAME, Optional ByVal timeout As ULong = 20) As Boolean

        Static failCounter As Integer = 0

        SyncLock lock_kvaser_CAN_TX

            If canHandle >= 0 Then

                Dim flag As Integer

                If Canlib.canReadStatus(canHandle, flag) = Canlib.canStatus.canOK Then

                    If (flag And Canlib.canSTAT_HW_OVERRUN) <> 0 Or (flag And Canlib.canSTAT_ERROR_PASSIVE) <> 0 Then
                        Canlib.canResetBus(canHandle)
                        Canlib.canBusOn(canHandle)
                    End If

                    Dim flag2 As Integer = 0

                    If cand.rtr Then
                        flag2 += Canlib.canMSG_RTR
                    End If
                    If cand.ide Then
                        flag2 += Canlib.canMSG_EXT
                    Else
                        flag2 += Canlib.canMSG_STD
                    End If

                    If Canlib.canWrite(canHandle, cand.id, cand.data, cand.data.Length, flag2) = Canlib.canStatus.canOK Then

                        If Canlib.canWriteSync(canHandle, timeout) = Canlib.canStatus.canOK Then
                            failCounter = 0
                            Return True
                        End If

                        Canlib.canResetBus(canHandle)
                        Canlib.canBusOn(canHandle)
                        failCounter += 1

                        If failCounter > 10 Then
                            'Throw New Exception("kvaser canTx error")
                        End If

                    End If
                End If
            End If

        End SyncLock

        Return False

    End Function


    ''' <summary>
    ''' フィルタ設定
    ''' </summary>
    ''' <param name="id">受信ID配列</param>
    ''' <param name="mask"></param>
    ''' <param name="ide"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function kvaserFilterSetting(ByVal id As Integer, ByVal mask As Integer, ByVal ide As Boolean) As Boolean

        If Canlib.canSetAcceptanceFilter(canHandle, id, mask, ide) = Canlib.canStatus.canOK Then
            Return True
        End If
        Return False

    End Function


    ''' <summary>
    ''' ウィンドウメッセージ処理（CANイベントを監視）
    ''' </summary>
    ''' <param name="m"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function WndMsgProc(ByRef m As System.Windows.Forms.Message) As Boolean
        Select Case m.Msg
            Case Canlib.WM__CANLIB
                Rx()
        End Select
        Return True
    End Function


    ''' <summary>
    ''' カバサーからoncommイベントを発生する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Rx()

        Dim args As New myEventArgs
        Dim stat As Long = 0

        If Canlib.canReadStatus(Me.canHandle, stat) = Canlib.canStatus.canOK Then

            If (stat And Canlib.canSTAT_HW_OVERRUN) <> 0 Or (stat And Canlib.canSTAT_ERROR_PASSIVE) <> 0 Then
                Canlib.canFlushReceiveQueue(Me.canHandle)       '受信バッファを消す
                Canlib.canResetBus(Me.canHandle)                'バスリセット
            End If

            'CANデータ受信処理
            Dim timestamp As Long   'タイムスタンプ
            Dim dlc As Integer      'データ受信数
            Dim flag As Integer     'フレーム形式
            Dim recvadta(7) As Byte '受信データ
            Dim recvid As Integer   '受信ID

            Do While Canlib.canRead(canHandle, recvid, recvadta, dlc, flag, timestamp) = Canlib.canStatus.canOK

                Dim candf As New CANDFRAME(recvid, recvadta, flag And &H1, flag And &H4, timestamp)

                candf.dlc = dlc 'DLC設定
                args.data = candf

                If Me.specialID.Contains(candf.id) Then
                    rcv_can_data.Copy(recvid, recvadta, candf.ide, candf.rtr, timestamp)
                    TxRx_Rcv = True
                End If

                RaiseEvent OnComm(Me, args)

                recvadta = New Byte(7) {}

            Loop

        End If

    End Sub




    Public Function writePacket(ByVal id As Integer, ByVal data() As Byte) As Boolean


    End Function


End Class


