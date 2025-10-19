using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ThwLib
{
    // ****************************************************************************************************
    /// <summary>
    /// Turbo HAMLOG/Win（THW）入力・修正ウィンドウ アクセスクラス<br/>
    /// Author: 7L2VKS
    /// </summary>
    /// <remarks>
    /// <code>
    /// 使い方：
    /// Thw thw = Thw.GetInstance();               // インスタンスの作成
    /// thw.TargetWindw = Thw.WindowType.Input;    // 修正ウィンドウに対する操作の場合は.Editを指定
    /// IntPtr hWnd = thw.SomeMethod();            // 機能の呼び出し
    /// ...
    /// </code>
    /// </remarks>
    // ****************************************************************************************************
    public class Thw : Form
    {
        [StructLayout(LayoutKind.Explicit)]
        private struct COPYDATASTRUCT
        {
            [FieldOffset(0)] public UInt64 dwData;  // コマンド
            [FieldOffset(8)] public UInt64 cbData;  // lpDataのバイト数
            [FieldOffset(16)] public IntPtr lpData; // 送信するデータへのポインタ
        }

        private const Int32 WM_COPYDATA = 0x004A;

        // ウィンドウ検索
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(String? lpClassName, String? lpWindowName);

        // メッセージ送信 (Send)
        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern Int32 SendMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, ref COPYDATASTRUCT lParam);

        // メッセージ送信 (Post)
        [DllImport("User32.dll", EntryPoint = "PostMessage")]
        private static extern Int32 PostMessage(IntPtr hWnd, Int32 Msg, Int32 wParam, ref COPYDATASTRUCT lParam);

        // フォーカス設定
        [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        private static extern IntPtr SetForegroundWindow(IntPtr hWnd);

        // THW class name & window title
        private const string THW_CLASS_NAME = "TThwin";
        private const string THW_EDITWINDOW_TITLE = "データの表示＆修正";

        // THW operational options
        private const UInt64 OP_ENTER = 0x00010000;         // データ送信後、ENTERキーを押下（Command=1～15）
        private const UInt64 OP_FOCUS = 0x00020000;         // データ送信後、送信対象項目にフォーカスを移動（Command=1～15）
        private const UInt64 OP_SAVEBOX_ON = 0x00040000;    // データ登録時、環境設定にかかわらず確認メッセージを表示（Command=18）
        private const UInt64 OP_SAVEBOX_OFF = 0x00080000;   // データ登録時、環境設定にかかわらず確認メッセージを非表示（Command=18）
        //private const UInt64 OP_APPLIHWND = 0x00100000;     (Not Support) メインウインドウのハンドルを返す（Command=101～117）
        //private const UInt64 OP_RMKS_RIGHT = 0x00100000;    (Not Support) 既存のRemarks文字列の右側に文字列を追加（Command=13～15）
        //private const UInt64 OP_RMKS_LEFT = 0x01000000;     (Not Support) 既存のRemarks文字列の左側に文字列を挿入（Command=13～15）

        // OP_SHUUSEI_WIN、OP_SHUUSEIWIN4を指定すると修正ウィンドウが操作の対象になる。ただし修正ウィンドウ
        // がフォアグラウンドにないと、入力ウィンドウに対して操作されてしまうという謎仕様。
        //private const UInt64 OP_SHUUSEI_WIN = 0x00200000;   (Not Support) 修正ウインドウへの送信データ書き込み確認を表示(v5.09a)
        private const UInt64 OP_SHUUSEIWIN4 = 0x00400000;   // 修正ウインドウから取り込む(v5.09a)

        private const UInt64 OP_QTH_MST = 0x00800000;       // QTHフィールドの内容はMSTのQTHを利用する（Command=15）
        //private const UInt64 OP_DX_ENTITY = 0x01000000;     (Not Support) DXエンティティ選択ウインドウを表示させる(v5.27c進捗版)

        // THW commands
        private const UInt64 CMD_SET_CALL = 1;
        private const UInt64 CMD_SET_DATE = 2;
        private const UInt64 CMD_SET_TIME = 3;
        private const UInt64 CMD_SET_HIS = 4;
        private const UInt64 CMD_SET_MY = 5;
        private const UInt64 CMD_SET_FREQ = 6;
        private const UInt64 CMD_SET_MODE = 7;
        private const UInt64 CMD_SET_CODE = 8;
        private const UInt64 CMD_SET_GL = 9;
        private const UInt64 CMD_SET_QSL = 10;
        private const UInt64 CMD_SET_HISNAME = 11;
        private const UInt64 CMD_SET_QTH = 12;
        private const UInt64 CMD_SET_REMARKS1 = 13;
        private const UInt64 CMD_SET_REMARKS2 = 14;
        private const UInt64 CMD_SET_ALL_INPUT = 15;
        private const UInt64 CMD_CLEAR = 16;
        private const UInt64 CMD_DUPCHECK = 17;
        private const UInt64 CMD_SAVE = 18;
        //private const UInt64 CMD_SHOWQTH = 19;              (Not Support) Codeフィールドで↓キー押下
        private const UInt64 CMD_SHOW_RECONEDIT = 20;
        private const UInt64 CMD_SET_ALL_EDIT = 21;
        //private const UInt64 CMD_PTT = 22;                  (Not Support) THWからPTTon/offコマンドを送出（toggle）
        //private const UInt64 CMD_PTT_ON = 23;               (Not Support) THWからPTTonコマンドを送出
        //private const UInt64 CMD_PTT_OFF = 24;              (Not Support) THWからPTToffコマンドを送出
        private const UInt64 CMD_CHECKBOX_ON = 25;
        private const UInt64 CMD_CHECKBOX_OFF = 26;
        //private const UInt64 CMD_QSODATA_CLOSE = 27;        (Not Support) QSOデータのクローズ
        //private const UInt64 CMD_QSODATA_OPEN = 28;         (Not Support) QSOデータのオープン
        private const UInt64 CMD_MISCELLANEOUS = 30;

        private const UInt64 CMD_GET_CALL = 101;
        private const UInt64 CMD_GET_DATE = 102;
        private const UInt64 CMD_GET_TIME = 103;
        private const UInt64 CMD_GET_HIS = 104;
        private const UInt64 CMD_GET_MY = 105;
        private const UInt64 CMD_GET_FREQ = 106;
        private const UInt64 CMD_GET_MODE = 107;
        private const UInt64 CMD_GET_CODE = 108;
        private const UInt64 CMD_GET_GL = 109;
        private const UInt64 CMD_GET_QSL = 110;
        private const UInt64 CMD_GET_HISNAME = 111;
        private const UInt64 CMD_GET_QTH = 112;
        private const UInt64 CMD_GET_REMARKS1 = 113;
        private const UInt64 CMD_GET_REMARKS2 = 114;
        private const UInt64 CMD_GET_ALL = 115;
        private const UInt64 CMD_GET_HDB_ATTRIBUTE = 116;
        private const UInt64 CMD_GET_DIRECTION = 117;
        private const UInt64 CMD_GET_ALL_AT = 118;
        private const UInt64 CMD_GET_ALL_FOR = 119;

        // THW checkbox flags
        private const int CHECKBOX_DX = 8;
        private const int CHECKBOX_CQ = 16;
        private const int CHECKBOX_1 = 32;
        private const int CHECKBOX_2 = 64;

        //　Event

        public delegate void SaveEventHandler(object sender, ThwSaveEventArgs e);
        /// <value>Give Upタイマータイムアウト後のSave処理完了イベント</value>
        public event SaveEventHandler? SaveEvent;

        // Properties

        /// <value>操作対象ウィンドウ (<c>WindowType.Input</c> - 入力ウィンドウ、 <c>WindowType.Edit</c> - 修正ウィンドウ)</value>
        public WindowType TargetWindow { get; set; } = WindowType.Input;

        /// <value>Give Upタイマー（秒）</value>
        public int GiveupTime { get; set; } = 0;

        /// <value>Save処理 (<c>true</c> - 処理中、 <c>false</c> - 非処理中)</value>
        public bool RunningSave { get; private set; } = false;

        // Fields
        private readonly Encoding _encoding;
        private string _message = string.Empty;
        private readonly IntPtr _hWndThis;
        private CancellationTokenSource? _ctsTimeout = null;
        private ExceptionDispatchInfo? _exceptionDispatch = null;

        // Constractor
        private Thw()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            _encoding = Encoding.GetEncoding("Shift-JIS");

            _hWndThis = this.Handle;
        }

        // ------------------------------------------------------------------------------------------------
        //   Internal Methods
        // ------------------------------------------------------------------------------------------------

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            switch (m.Msg)
            {
                case WM_COPYDATA:
                    COPYDATASTRUCT cds = (COPYDATASTRUCT)Marshal.PtrToStructure(m.LParam, typeof(COPYDATASTRUCT))!;
                    int length = (int)cds.cbData - 1;       // Do not include last \0 character
                    var shiftjisBytes = new byte[length / Marshal.SizeOf(typeof(byte))];

                    // メモリ領域の値を配列に複写
                    Marshal.Copy(cds.lpData, shiftjisBytes, 0, shiftjisBytes.Length);
                    _message = _encoding.GetString(shiftjisBytes);

                    break;
            }
        }

        private IntPtr SendMessage(ref COPYDATASTRUCT cds)
        {
            IntPtr hWndThw = FindWindow(THW_CLASS_NAME, null);
            if (hWndThw == IntPtr.Zero)
            {
                throw new ThwException(ThwError.THWNotFound);
            }

            //IntPtr hWnd = (IntPtr)SendMessage(hWndThw, WM_COPYDATA, this.Handle, ref cds);
            IntPtr hWnd = (IntPtr)SendMessage(hWndThw, WM_COPYDATA, _hWndThis, ref cds);
            if (hWnd == IntPtr.Zero)
            {
                throw new ThwException(ThwError.InvalidWindowHandle);
            }

            return hWnd;
        }

        private IntPtr SendCommand(UInt64 command)
        {
            COPYDATASTRUCT cds;

            cds.dwData = command;
            cds.cbData = 0;
            cds.lpData = IntPtr.Zero;

            return SendMessage(ref cds);
        }

        private IntPtr SendCommand(UInt64 command, string str)
        {
            IntPtr hWnd;
            COPYDATASTRUCT cds;

            cds.dwData = command;

            byte[] shiftjisBytes = _encoding.GetBytes(str);

            int allocateSize = Marshal.SizeOf(typeof(byte)) * shiftjisBytes.Length;
            // Unmanagedメモリ領域を確保
            IntPtr pUnmanaged = Marshal.AllocHGlobal(allocateSize);
            // メモリ領域に配列の値を複写
            Marshal.Copy(shiftjisBytes, 0, pUnmanaged, shiftjisBytes.Length);

            try
            {
                cds.cbData = (UInt64)allocateSize;
                cds.lpData = pUnmanaged;
                hWnd = SendMessage(ref cds);
            }
            finally
            {
                // Unmanagedメモリ領域の解放 Do not forget it !!!
                Marshal.FreeHGlobal(pUnmanaged);
            }

            return hWnd;
        }

        private IntPtr SetData(UInt64 command, string str, bool setFocus, bool pressEnter)
            => SendCommand(command + (setFocus ? OP_FOCUS : 0) + (pressEnter ? OP_ENTER : 0), str);

        private IntPtr GetData(UInt64 command, out string outStr)
        {
            IntPtr hWnd = SendCommand(command);
            outStr = _message;

            return hWnd;
        }

        private IntPtr GetData(UInt64 command, string str, out string outStr)
        {
            IntPtr hWnd = SendCommand(command, str);
            outStr = _message;

            return hWnd;
        }

        private UInt64 SelectCommand(UInt64 command)
        {
            if (TargetWindow == WindowType.Edit)
            {
                // 修正ウィンドウに対する操作を行う場合、そのウィンドウがフォアグラウンドにないと、
                // 入力ウィンドウに対して操作されてしまうため、あらかじめフォアグラウンドに移動して
                // おく。
                IntPtr hWnd = FindWindow(null, THW_EDITWINDOW_TITLE);
                if (hWnd == IntPtr.Zero)
                {
                    throw new ThwException(ThwError.EditWindowNotFound);
                }
                SetForegroundWindow(hWnd);

                command |= OP_SHUUSEIWIN4;
            }

            return command;
        }

        // ------------------------------------------------------------------------------------------------
        //   入力・操作系
        // ------------------------------------------------------------------------------------------------

        /// <summary>
        /// 文字列入力（Callフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetCall(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_CALL), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Dateフィールド）
        /// </summary>
        /// <param name="date">入力日付</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetDate(DateTime date, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_DATE), date.ToString("yy/MM/dd"), setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Timeフィールド）
        /// </summary>
        /// <remarks>
        /// <paramref name="time"/>パラメータのKindプロパティにDateTimeKind.Utcが設定されている場合はUTC、それ以外はJSTとして入力される。
        /// </remarks>
        /// <param name="time">入力時刻</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetTime(DateTime time, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_TIME), time.ToString("HH:mm") + (time.Kind == DateTimeKind.Utc ? "U" : "J"), setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Date/Timeフィールド）
        /// </summary>
        /// <remarks>
        /// <paramref name="dateTime"/>パラメータのKindプロパティにDateTimeKind.Utcが設定されている場合はUTC、それ以外はJSTとして入力される。
        /// </remarks>
        /// <param name="dateTime">入力日時</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetDateTime(DateTime dateTime, bool setFocus, bool pressEnter)
        {
            SetDate(dateTime, setFocus, pressEnter);
            return SetTime(dateTime, setFocus, pressEnter);
        }

        /// <summary>
        /// 文字列入力（Hisフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetHis(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_HIS), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Myフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetMy(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_MY), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Freqフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetFreq(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_FREQ), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Modeフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetMode(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_MODE), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Codeフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetCode(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_CODE), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（GLフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetGL(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_GL), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（QSLフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetQSL(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_QSL), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（QSLフィールド）
        /// </summary>
        /// <remarks>
        /// 指定した入力位置に空文字を入力する場合は、<paramref name="ch"/>にnullを設定する。
        /// </remarks>
        /// <param name="ch">入力文字またはnull</param>
        /// <param name="position">入力位置（左から0, 1, 2）</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetQSL(char? ch, int position, bool setFocus, bool pressEnter)
        {
            if (position < 0 || position > 2)
            {
                throw new ThwException(ThwError.InvalidArgument, "position");
            }

            GetQSL(out string str);

            str = string.Concat(str, "   ");
            char[] qslChars = str.ToCharArray();
            qslChars[position] = ch ?? ' ';

            return SetData(SelectCommand(CMD_SET_QSL), new string(qslChars).TrimEnd(), setFocus, pressEnter);
        }

        /// <summary>
        /// 文字列入力（His Nameフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetHisName(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_HISNAME), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（QTHフィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetQTH(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_QTH), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Remarks1フィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetRemarks1(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_REMARKS1), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（Remarks2フィールド）
        /// </summary>
        /// <param name="str">入力文字列</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetRemarks2(string str, bool setFocus, bool pressEnter)
            => SetData(SelectCommand(CMD_SET_REMARKS2), str, setFocus, pressEnter);

        /// <summary>
        /// 文字列入力（全フィールド）
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>空文字列 ""（DateTimeはnull）を指定するとそのフィールドはスキップされる（クリアはされない）。</item>
        /// <item>codeに正しいコードが指定され、qthが空文字列の場合は、QTHフィールドの内容はMSTのQTHを利用する。</item>
        /// <item>対象が修正ウィンドウの場合は、Callフィールドの変更は不可。</item>
        /// </list>
        /// </remarks>
        /// <param name="qso">入力データを指定したThwQSORecordインスタンスへの参照</param>
        /// <param name="setFocus"><c>true</c> - 入力対象項目のフィールドにフォーカスを移動</param>
        /// <param name="pressEnter"><c>true</c> - 処理後、ENTERキーを押下</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetAll(ThwQSORecord qso, bool setFocus, bool pressEnter)
        {
            ThwQSORecord.DateTimeToString(qso.DateTime, out string date, out string time);
            string checkBoxes = Convert.ToString((qso.CheckDX ? CHECKBOX_DX : 0)
                                               + (qso.CheckCQ ? CHECKBOX_CQ : 0)
                                               + (qso.Check1 ? CHECKBOX_1 : 0)
                                               + (qso.Check2 ? CHECKBOX_2 : 0));
            return SetData(
                (TargetWindow == WindowType.Input ? CMD_SET_ALL_INPUT : CMD_SET_ALL_EDIT) | (qso.Code.Length != 0 && qso.QTH.Length == 0 ? OP_QTH_MST : 0),
                string.Join($"{Environment.NewLine}",
                            string.Empty,
                            qso.Call,
                            date,
                            time,
                            qso.His,
                            qso.My,
                            qso.Freq,
                            qso.Mode,
                            qso.Code,
                            qso.GL,
                            qso.QSL,
                            qso.HisName,
                            qso.QTH,
                            qso.Remarks1,
                            qso.Remarks2,
                            checkBoxes) + $"{Environment.NewLine}",
                setFocus, pressEnter);
        }

        /// <summary>
        /// 入力ウィンドウ：入力バッファクリア（Clearボタン押下）<br/>
        /// 修正ウィンドウ：修正の取り消し（ESCキー押下）
        /// </summary>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Clear() => SendCommand(SelectCommand(CMD_CLEAR));

        /// <summary>
        /// デュープチェック（F3キー押下）
        /// </summary>
        /// <returns>交信履歴ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr DupCheck() => SendCommand(SelectCommand(CMD_DUPCHECK));

        /// <summary>
        /// コールサイン入力＋デュープチェック（F3キー押下）
        /// </summary>
        /// <param name="call">Callフィールドへの入力文字列</param>
        /// <param name="setFocus"><c>true</c> - Callフィールドにフォーカスを移動</param>
        /// <returns>交信履歴ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr DupCheck(string call, bool setFocus)
        {
            SetCall(call, setFocus, false);
            return DupCheck();
        }

        /// <summary>
        /// データ登録（Saveボタン押下）
        /// </summary>
        /// <remarks>
        /// 入力ウィンドウではコールサイン入力＋Enterキー押下をしておかないとデータ登録が実行されないため（THWの仕様）、
        /// 事前に以下を呼び出す必要がある。<br/>
        /// <c>SetCall(call = "callsign", setFocus = true, pressEnter = true);</c>
        /// </remarks>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Save() => SendCommand(SelectCommand(CMD_SAVE));

        /// <summary>
        /// データ登録（入力ウィンドウ Saveボタン押下）
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>修正ウィンドウではサポートされない。</item>
        /// <item>コールサイン入力＋Enterキー押下をしておかないとデータ登録が実行されないため（THWの仕様）、事前に以下を呼び出す必要がある。<br/>
        /// <c>SetCall(call = "callsign", setFocus = true, pressEnter = true);</c></item>
        /// </list>
        /// </remarks>
        /// <param name="confirm">環境設定にかかわらず確認メッセージを <c>true</c> - 表示、 <c>false</c> - 非表示</param>
        /// <returns>入力ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Save(bool confirm)
        {
            if (TargetWindow == WindowType.Edit)
            {
                throw new ThwException(ThwError.NotSupportedFunction);
            }
            return SendCommand(CMD_SAVE + (confirm ? OP_SAVEBOX_ON : OP_SAVEBOX_OFF));
        }

        /// <summary>
        /// データ登録（Saveボタン押下）（非同期）
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>入力ウィンドウではコールサイン入力＋Enterキー押下をしておかないとデータ登録が実行されないため（THWの仕様）、
        /// 事前に以下を呼び出す必要がある。<br/>
        /// <c>SetCall(call = "callsign", setFocus = true, pressEnter = true);</c></item>
        /// <item>UIを持つアプリケーションでは、登録時にTHW上で確認ダイアログが表示されるとそれに応答するまで制御が戻らないため、
        /// UIがロックされる。その場合はこの非同期版を使用する。</item>
        /// <item>Taskは以下のいずれかの条件を満たすと終了する。
        /// <list type="bullet">
        /// <item>THWが制御を戻した時。</item>
        /// <item>Give Upタイマーが設定されている場合、タイマーがタイムアップした時。</item>
        /// </list>
        /// </list>
        /// </remarks>
        /// <returns>入力または修正ウィンドウのHandleをResultとするTask</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public Task<IntPtr> SaveAsync()
        {
            if (RunningSave)
            {
                throw new ThwException(ThwError.FunctionAlreadyRunning);
            }
            return SaveAsync(SelectCommand(CMD_SAVE));
        }

        /// <summary>
        /// データ登録（入力ウィンドウ Saveボタン押下）（非同期）
        /// </summary>
        /// <remarks>
        /// <list type="bullet">
        /// <item>修正ウィンドウではサポートされない。</item>
        /// <item>コールサイン入力＋Enterキー押下をしておかないとデータ登録が実行されないため（THWの仕様）、事前に以下を呼び出す必要がある。<br/>
        /// <c>SetCall(call = "callsign", setFocus = true, pressEnter = true);</c></item>
        /// <item>UIを持つアプリケーションでは、登録時にTHW上で確認ダイアログが表示されるとそれに応答するまで制御が戻らないため、
        /// UIがロックされる。その場合はこの非同期版を使用する。</item>
        /// <item>Taskは以下のいずれかの条件を満たすと終了する。
        /// <list type="bullet">
        /// <item>THWが制御を戻した時。</item>
        /// <item>Give Upタイマーが設定されている場合、タイマーがタイムアップした時。</item>
        /// </list>
        /// </list>
        /// </remarks>
        /// <param name="confirm">環境設定にかかわらず確認メッセージを <c>true</c> - 表示、 <c>false</c> - 非表示</param>
        /// <returns>入力ウィンドウのHandleをResultとするTask</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public Task<IntPtr> SaveAsync(bool confirm)
        {
            if (TargetWindow == WindowType.Edit)
            {
                throw new ThwException(ThwError.NotSupportedFunction);
            }
            if (RunningSave)
            {
                throw new ThwException(ThwError.FunctionAlreadyRunning);
            }
            return SaveAsync(CMD_SAVE + (confirm ? OP_SAVEBOX_ON : OP_SAVEBOX_OFF));
        }

        private async Task<IntPtr> SaveAsync(UInt64 command)
        {
            IntPtr hWnd;

            Task<IntPtr> task = RunSaveAsync(command);
            using (_ctsTimeout = new CancellationTokenSource())
            {
                try
                {
                    await Task.Delay(GiveupTime != 0 ? GiveupTime * 1000 : -1, _ctsTimeout.Token);

                    // Save処理が完了せず、Delay()がキャンセルされなかった（Save処理完了時にRunSaveAsync()内でSaveEventイベント発生）
                    throw new ThwException(ThwError.RequestTimeout);
                }
                catch (OperationCanceledException)
                {
                    // Save処理が完了し、Delay()がキャンセルされた
                    // ・task.Resultがthrowする例外はAggregateExceptionとなるので、ExceptionDispatchInfoを使用している
                    // ・RunSaveAsync()内でExceptionをcatchしているので、task.Resultから例外がthrowされることはない
                    hWnd = task.Result;
                    _exceptionDispatch?.Throw();
                }
                finally
                {
                    _ctsTimeout = null;
                }
            }

            return hWnd;
        }

        private async Task<IntPtr> RunSaveAsync(UInt64 command)
        {
            Exception? exception = null;
            IntPtr hWnd = IntPtr.Zero;

            _exceptionDispatch = null;

            try
            {
                RunningSave = true;
                hWnd = await Task.Run(() => SendCommand(command));
            }
            catch (Exception ex)
            {
                exception = ex;
                _exceptionDispatch = ExceptionDispatchInfo.Capture(ex);
            }
            RunningSave = false;

            if (_ctsTimeout == null)
            {
                ThwSaveEventArgs e = new() { Result = hWnd != IntPtr.Zero, HWnd = hWnd, Exception = exception };
                SaveEvent?.Invoke(this, e);
            }
            else
            {
                _ctsTimeout.Cancel();
            }

            return hWnd;
        }

        /// <summary>
        /// 修正ウィンドウ表示
        /// </summary>
        /// <remarks>
        /// <paramref name="recNo"/>パラメータに指定したレコード番号が範囲外の場合、最後のレコードが対象となる。
        /// </remarks>
        /// <param name="recNo">修正ウィンドウに表示するレコードの番号</param>
        /// <returns>修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr ShowRecOnEdit(int recNo) => SendCommand(CMD_SHOW_RECONEDIT, Convert.ToString(recNo));

        /// <summary>
        /// 「CQ」チェックボックス設定
        /// </summary>
        /// <param name="on"><c>true</c> - オン、 <c>false</c> - オフ</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetCQ(bool on) => SendCommand(SelectCommand(on ? CMD_CHECKBOX_ON : CMD_CHECKBOX_OFF), "1");

        /// <summary>
        /// 「1」チェックボックス設定
        /// </summary>
        /// <param name="on"><c>true</c> - オン、 <c>false</c> - オフ</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Set1(bool on) => SendCommand(SelectCommand(on ? CMD_CHECKBOX_ON : CMD_CHECKBOX_OFF), "2");

        /// <summary>
        /// 「2」チェックボックス設定
        /// </summary>
        /// <param name="on"><c>true</c> - オン、 <c>false</c> - オフ</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Set2(bool on) => SendCommand(SelectCommand(on ? CMD_CHECKBOX_ON : CMD_CHECKBOX_OFF), "3");

        /// <summary>
        /// 「DX」チェックボックス設定
        /// </summary>
        /// <param name="on"><c>true</c> - オン、 <c>false</c> - オフ</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetDX(bool on) => SendCommand(SelectCommand(on ? CMD_CHECKBOX_ON : CMD_CHECKBOX_OFF), "4");

        /// <summary>
        /// DXエンティティ自動入力モード設定
        /// </summary>
        /// <remarks>
        /// デフォルトは自動入力モードオン。
        /// </remarks>
        /// <param name="on">Code/QTHフィールドへのDXエンティティ自動入力モード <c>true</c> - オン、 <c>false</c> - オフ</param>
        /// <returns>入力ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr SetAutoDXEntity(bool on) => SendCommand(CMD_MISCELLANEOUS, on ? "2" : "1");

        /// <summary>
        /// アプリサイズの最小化
        /// </summary>
        /// <returns>修正のウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Minimize() => SendCommand(CMD_MISCELLANEOUS, "3");

        /// <summary>
        /// アプリサイズのリストア
        /// </summary>
        /// <returns>修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr Restore() => SendCommand(CMD_MISCELLANEOUS, "4");

        // ------------------------------------------------------------------------------------------------
        //   取得系
        // ------------------------------------------------------------------------------------------------

        /// <summary>
        /// 文字列取得（Callフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetCall(out string str) => GetData(SelectCommand(CMD_GET_CALL), out str);

        /// <summary>
        /// 文字列取得（Dateフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetDate(out string str) => GetData(SelectCommand(CMD_GET_DATE), out str);

        /// <summary>
        /// DateTimeオブジェクト取得（Dateフィールド）
        /// </summary>
        /// <remarks>
        /// <paramref name="date"/>にDateTimeインスタンスへの参照が返された場合、時刻は0:00:00、
        /// KindプロパティはDateTimeKind.Unspecifiedが設定される。
        /// </remarks>
        /// <param name="date">作成されたDateTimeインスタンスへの参照。無効な日付文字列が指定された場合はnull。</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetDate(out DateTime? date)
        {
            IntPtr hWnd = GetDate(out string str);
            date = ThwQSORecord.StringToDate(str);
            return hWnd;
        }

        /// <summary>
        /// 文字列取得（Timeフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetTime(out string str) => GetData(SelectCommand(CMD_GET_TIME), out str);

        /// <summary>
        /// DateTimeオブジェクト取得（Timeフィールド）
        /// </summary>
        /// <remarks>
        /// <paramref name="time"/>にDateTimeインスタンスへの参照が返された場合、日付は0001-01-01が設定される。
        /// </remarks>
        /// <param name="time">作成されたDateTimeインスタンスへの参照。無効な時刻文字列が指定された場合はnull。</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetTime(out DateTime? time)
        {
            IntPtr hWnd = GetTime(out string str);
            time = ThwQSORecord.StringToTime(str);

            return hWnd;
        }

        /// <summary>
        /// 文字列取得（Date/Timeフィールド）
        /// </summary>
        /// <param name="date">Dateフィールドからの取得文字列</param>
        /// <param name="time">Timeフィールドからの取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetDateTime(out string date, out string time)
        {
            IntPtr hWnd;
            hWnd = GetDate(out date); hWnd = GetTime(out time);

            return hWnd;
        }

        /// <summary>
        /// DateTimeオブジェクト取得（Date/Timeフィールド）
        /// </summary>
        /// <param name="dateTime">作成されたDateTimeインスタンスへの参照。無効な日付または時刻文字列が指定された場合はnull。</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetDateTime(out DateTime? dateTime)
        {
            IntPtr hWnd;

            hWnd = GetDateTime(out string date, out string time);
            dateTime = ThwQSORecord.StringToDateTime(date, time);

            return hWnd;
        }

        /// <summary>
        /// DateTimeオブジェクト取得（Date/Timeフィールド）
        /// </summary>
        /// <param name="dateTime">作成されたDateTimeインスタンスへの参照</param>
        /// <param name="timeZone"><c>TimeZone.JST</c> - UTCの場合JSTに変換、 <c>TimeZone.UTC</c> - JSTの場合UTCに変換</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetDateTime(out DateTime? dateTime, TimeZone timeZone)
        {
            IntPtr hWnd = GetDateTime(out dateTime);
            dateTime = timeZone == TimeZone.JST ? dateTime?.ToLocalTime() : dateTime?.ToUniversalTime();

            return hWnd;
        }

        /// <summary>
        /// 文字列取得（Hisフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetHis(out string str) => GetData(SelectCommand(CMD_GET_HIS), out str);

        /// <summary>
        /// 文字列取得（Myフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetMy(out string str) => GetData(SelectCommand(CMD_GET_MY), out str);

        /// <summary>
        /// 文字列取得（Freqフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetFreq(out string str) => GetData(SelectCommand(CMD_GET_FREQ), out str);

        /// <summary>
        /// 文字列取得（Modeフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetMode(out string str) => GetData(SelectCommand(CMD_GET_MODE), out str);

        /// <summary>
        /// 文字列取得（Codeフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetCode(out string str) => GetData(SelectCommand(CMD_GET_CODE), out str);

        /// <summary>
        /// 文字列取得（GLフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetGL(out string str) => GetData(SelectCommand(CMD_GET_GL), out str);

        /// <summary>
        /// 文字列取得（QSLフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetQSL(out string str) => GetData(SelectCommand(CMD_GET_QSL), out str);

        /// <summary>
        /// 文字取得（QSLフィールド）
        /// </summary>
        /// <remarks>
        /// 指定した取得位置に文字の入力がない場合は、<paramref name="ch"/>にnullが設定される。
        /// </remarks>
        /// <param name="ch">取得文字またはnull</param>
        /// <param name="position">取得位置（左から0, 1, 2）</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetQSL(out char? ch, int position)
        {
            if (position < 0 || position > 2)
            {
                throw new ThwException(ThwError.InvalidArgument, "position");
            }
            IntPtr hWnd = GetData(SelectCommand(CMD_GET_QSL), out string str);
            ch = str.Length >= position + 1 ? str[position] : null;

            return hWnd;
        }

        /// <summary>
        /// 文字列取得（His Nameフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetHisName(out string str) => GetData(SelectCommand(CMD_GET_HISNAME), out str);

        /// <summary>
        /// 文字列取得（QTHフィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetQTH(out string str) => GetData(SelectCommand(CMD_GET_QTH), out str);

        /// <summary>
        /// 文字列取得（Remarks1フィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetRemarks1(out string str) => GetData(SelectCommand(CMD_GET_REMARKS1), out str);

        /// <summary>
        /// 文字列取得（Remarks2フィールド）
        /// </summary>
        /// <param name="str">取得文字列</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetRemarks2(out string str) => GetData(SelectCommand(CMD_GET_REMARKS2), out str);

        /// <summary>
        /// 文字列取得（全フィールド）
        /// </summary>
        /// <param name="qso">取得データをコピーするThwQSORecordインスタンスへの参照</param>
        /// <returns>入力または修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetAll(ThwQSORecord qso)
        {
            IntPtr hWnd = GetData(SelectCommand(CMD_GET_ALL), out string str);
            ParseData(str, qso);

            return hWnd;
        }

        /// <summary>
        /// 文字列取得（選択レコード全フィールド）
        /// </summary>
        /// <remarks>
        /// メインウィンドウで選択されているレコードのデータを取得する。
        /// </remarks>
        /// <param name="qso">取得データをコピーするThwQSORecordインスタンスへの参照</param>
        /// <returns>修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetAllAt(ThwQSORecord qso)
        {
            IntPtr hWnd = GetData(CMD_GET_ALL_AT, out string str);
            ParseData(str, qso);

            return hWnd;
        }

        /// <summary>
        /// 文字列取得（指定レコード全フィールド）
        /// </summary>
        /// <remarks>
        /// <paramref name="recNo"/>で指定したレコード番号が範囲外の場合、最後のレコードが対象となる。
        /// </remarks>
        /// <param name="recNo">レコード番号</param>
        /// <param name="qso">取得データをコピーするThwQSORecordインスタンスへの参照</param>
        /// <returns>修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetAllAt(ThwQSORecord qso, int recNo)
        {
            IntPtr hWnd = GetData(CMD_GET_ALL_FOR, Convert.ToString(recNo), out string str);
            ParseData(str, qso);

            return hWnd;
        }

        private void ParseData(string str, ThwQSORecord qso)
        {
            string[] fields = str.Split($"{Environment.NewLine}", StringSplitOptions.None);

            qso.Call = fields[1];
            qso.DateTime = ThwQSORecord.StringToDateTime(fields[2], fields[3]);
            qso.His = fields[4];
            qso.My = fields[5];
            qso.Freq = fields[6];
            qso.Mode = fields[7];
            qso.Code = fields[8];
            qso.GL = fields[9];
            qso.QSL = fields[10];
            qso.HisName = fields[11];
            qso.QTH = fields[12];
            qso.Remarks1 = fields[13];
            qso.Remarks2 = fields[14];

            int flags = int.Parse(fields[15]);
            qso.CheckDX = (flags & CHECKBOX_DX) != 0;
            qso.CheckCQ = (flags & CHECKBOX_CQ) != 0;
            qso.Check1 = (flags & CHECKBOX_1) != 0;
            qso.Check2 = (flags & CHECKBOX_2) != 0;
        }

        /// <summary>
        /// HAMLOG.HDBファイルのフルパス・レコード件数取得
        /// </summary>
        /// <param name="path">HAMLOG.HDBファイルのフルパス</param>
        /// <param name="count">HAMLOG.HDBファイル内のレコード件数</param>
        /// <returns>修正ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetHDBInfo(out string path, out int count)
        {
            IntPtr hWnd = GetData(CMD_GET_HDB_ATTRIBUTE, out string str);

            var reg = new Regex(@$"Count=(\d+){Environment.NewLine}HDB=(.*){Environment.NewLine}");
            var m = reg.Match(str);

            if (m.Groups.Count >= 3)
            {
                count = int.Parse(m.Groups[1].Value);
                path = m.Groups[2].Value;
            }
            else
            {
                throw new ThwException(ThwError.InvalidArgument, str);
            }

            return hWnd;
        }

        /// <summary>
        /// 方位・距離の取得
        /// </summary>
        /// <remarks>
        /// 入力ウィンドウに方位・距離が表示されていない場合は、<paramref name="direction"/>および<paramref name="distance"/>に
        /// nullが設定される。
        /// </remarks>
        /// <param name="direction">方位（度）またはnull</param>
        /// <param name="distance">距離（km）またはnull</param>
        /// <returns>入力ウィンドウのHandle</returns>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public IntPtr GetDirction(out int? direction, out float? distance)
        {
            direction = null; distance = null;

            IntPtr hWnd = GetData(CMD_GET_DIRECTION, out string str);
            if (str.Length != 0)
            {
                var reg = new Regex(@"方位:(\d+)度\s+(\d+\.*\d*)km");
                var m = reg.Match(str);

                if (m.Groups.Count >= 3)
                {
                    direction = int.Parse(m.Groups[1].Value);
                    distance = float.Parse(m.Groups[2].Value);
                }
                else
                {
                    throw new ThwException(ThwError.InvalidArgument, str);
                }
            }

            return hWnd;
        }

        // ------------------------------------------------------------------------------------------------
        //   その他
        // ------------------------------------------------------------------------------------------------

        /// <summary>
        /// 指定ウィンドウを前面に表示
        /// </summary>
        /// <param name="hWnd">前面に表示するウィンドウのHandle</param>
        /// <exception cref="ThwException">THWアクセス・処理エラー</exception>
        public static void SetForeground(IntPtr hWnd)
        {
            SetForegroundWindow(hWnd);
        }

        /// <summary>
        /// Thwオブジェクトの作成
        /// </summary>
        /// <returns>Thwクラスのインスタンスへの参照</returns>
        public static Thw GetInstance()
        {
            var thw = new Thw() { Visible = false };
            return thw;
        }

        /// <summary>
        /// 操作対象ウィンドウ列挙子
        /// </summary>
        public enum WindowType
        {
            Input,
            Edit
        }

        /// <summary>
        /// タイムゾーン列挙子
        /// </summary>
        public enum TimeZone
        {
            JST,
            UTC
        }
    }

    // ****************************************************************************************************
    /// <summary>
    /// Turbo HAMLOG/Win（THW）QSOレコードクラス    
    　  /// </summary>
    // ****************************************************************************************************
    public class ThwQSORecord
    {
        /// <value>コールサイン</value>
        public string Call { get; set; } = string.Empty;

        /// <value>交信日時</value>
        public DateTime? DateTime { get; set; } = null;

        /// <value>レポート（送信）</value>
        public string His { get; set; } = string.Empty;

        /// <value>レポート（受信）</value>
        public string My { get; set; } = string.Empty;

        /// <value>周波数</value>
        public string Freq { get; set; } = string.Empty;

        /// <value>モード</value>
        public string Mode { get; set; } = string.Empty;

        /// <value>コード</value>
        public string Code { get; set; } = string.Empty;

        /// <value>グリッドロケータ</value>
        public string GL { get; set; } = string.Empty;

        /// <value>QSL交換履歴</value>
        public string QSL { get; set; } = string.Empty;

        /// <value>名前</value>
        public string HisName { get; set; } = string.Empty;

        /// <value>QTH</value>
        public string QTH { get; set; } = string.Empty;

        /// <value>Remarks 1</value>
        public string Remarks1 { get; set; } = string.Empty;

        /// <value>Remarks 2</value>
        public string Remarks2 { get; set; } = string.Empty;

        /// <value>DX チェックボックス</value>
        public bool CheckDX { get; set; } = false;

        /// <value>CQ チェックボックス</value>
        public bool CheckCQ { get; set; } = false;

        /// <value>1 チェックボックス</value>
        public bool Check1 { get; set; } = false;

        /// <value>2 チェックボックス</value>
        public bool Check2 { get; set; } = false;

        /// <summary>
        /// DateTimeから日付文字列へ変換
        /// </summary>
        /// <remarks>
        /// <paramref name="dateTime"/>がnullの場合は空文字列が設定される。
        /// </remarks>
        /// <param name="dateTime">DateTimeインスタンスへの参照またはnull</param>
        /// <returns>文字列へ変換された日付</returns>
        public static string DateToString(DateTime? dateTime)
            => dateTime?.ToString("yy/MM/dd") ?? string.Empty;

        /// <summary>
        /// DateTimeから時刻文字列へ変換
        /// </summary>
        /// <remarks>
        /// <paramref name="dateTime"/>がnullの場合は空文字列が設定される。
        /// </remarks>
        /// <param name="dateTime">DateTimeインスタンスへの参照またはnull</param>
        /// <returns>文字列へ変換された時刻</returns>
        public static string TimeToString(DateTime? dateTime)
            => dateTime != null ? (dateTime?.ToString("HH:mm") + (dateTime?.Kind == DateTimeKind.Utc ? "U" : "J")) : string.Empty;

        /// <summary>
        /// DateTimeから日付、時刻文字列へ変換
        /// </summary>
        /// <remarks>
        /// <paramref name="dateTime"/>がnullの場合は空文字列が設定される。
        /// </remarks>
        /// <param name="dateTime">DateTimeインスタンスへの参照またはnull</param>
        /// <param name="date">文字列へ変換された日付</param>
        /// <param name="time">文字列へ変換された時刻</param>
        public static void DateTimeToString(DateTime? dateTime, out string date, out string time)
        {
            date = DateToString(dateTime);
            time = TimeToString(dateTime);
        }

        /// <summary>
        /// 日付文字列からDateTimeへ変換
        /// </summary>
        /// <remarks>
        /// <paramref name="date"/>パラメータの時刻は0:00:00、KindプロパティはDateTimeKind.Unspecifiedが設定される。
        /// </remarks>
        /// <param name="date">日付文字列</param>
        /// <returns>作成されたDateTimeインスタンスへの参照。無効な日付文字列が指定された場合はnull。</returns>
        public static DateTime? StringToDate(string date)
        {
            bool result = ParseDate(date, out int year, out int month, out int day);
            return result ? new DateTime(year, month, day) : null;
        }

        /// <summary>
        /// 時刻文字列からDateTimeへ変換
        /// </summary>
        /// <remarks>
        /// <paramref name="time"/>パラメータの日付は0001-01-01が設定される。
        /// </remarks>
        /// <param name="time">時刻文字列</param>
        /// <returns>作成されたDateTimeインスタンスへの参照。無効な時刻文字列が指定された場合はnull。</returns>
        public static DateTime? StringToTime(string time)
        {
            bool result = ParseTime(time, out int hour, out int minute, out bool utc);
            return result ? new DateTime(1, 1, 1, hour, minute, 0, utc ? DateTimeKind.Utc : DateTimeKind.Local) : null;
        }

        /// <summary>
        /// 日時文字列からDateTimeへ変換
        /// </summary>
        /// <param name="date">日付文字列</param>
        /// <param name="time">時刻文字列</param>
        /// <returns>作成されたDateTimeインスタンスへの参照。無効な日付または時刻文字列が指定された場合はnull。</returns>
        public static DateTime? StringToDateTime(string date, string time)
        {
            bool result = ParseDate(date, out int year, out int month, out int day) // call both methods anyway to set value...
                        & ParseTime(time, out int hour, out int minute, out bool utc);
            return result ? new DateTime(year, month, day, hour, minute, 0, utc ? DateTimeKind.Utc : DateTimeKind.Local) : null;
        }

        private static bool ParseDate(string date, out int year, out int month, out int day)
        {
            year = month = day = 0;

            var reg = new Regex(@"(\d+)/(\d+)/(\d+)");
            var m = reg.Match(date);

            if (m.Groups.Count < 4)
                return false;

            year = int.Parse(m.Groups[1].Value);
            month = int.Parse(m.Groups[2].Value);
            day = int.Parse(m.Groups[3].Value);

            year += year < 50 ? 2000 : 1900;

            if (month > 12 || day > 31)
                return false;

            return true;
        }

        private static bool ParseTime(string time, out int hour, out int minute, out bool utc)
        {
            hour = minute = 0; utc = false;

            var reg = new Regex(@"(\d+):(\d+)(.)");
            var m = reg.Match(time);

            if (m.Groups.Count < 4)
                return false;

            hour = int.Parse(m.Groups[1].Value);
            minute = int.Parse(m.Groups[2].Value);
            utc = m.Groups[3].Value == "U";

            if (hour > 23 || minute > 59)
                return false;

            return true;
        }
    }

    // ****************************************************************************************************
    /// <summary>
    /// イベント引数クラス（Save処理完了）
    /// </summary>
    // ****************************************************************************************************
    public class ThwSaveEventArgs
    {
        /// <value>処理結果： <c>true</c> - 成功、 <c>false</c> - 失敗</value>
        public bool Result { get; set; }

        /// <value>入力または修正ウィンドウのHandle</value>
        public IntPtr HWnd { get; set; }

        /// <value>処理結果（例外）</value>
        public Exception? Exception { get; set; } = null;
    }

    // ****************************************************************************************************
    /// <summary>
    /// Thw例外クラス
    /// </summary>
    // ****************************************************************************************************
    public class ThwException : Exception
    {
        private static readonly Dictionary<ThwError, string> errorMessage = new()
        {
            { ThwError.THWNotFound,            "Turbo HAMLOG/Win is not started." },
            { ThwError.EditWindowNotFound,     "Edit window was not found." },
            { ThwError.InvalidWindowHandle,    "Invalid input window handle was returend (internal error)." },
            { ThwError.InvalidArgument,        "Invalid argument was specified." },
            { ThwError.NotSupportedFunction,   "This function is not supported in this status." },
            { ThwError.FunctionAlreadyRunning, "This function cannot be started because it is already running." },
            { ThwError.RequestTimeout,         "The request was not completed within the specified time." },
        };

        /// <value>例外種別</value>
        public ThwError Error { get; }

        public ThwException(ThwError error) : base(errorMessage[error])
        {
            Error = error;
        }

        public ThwException(ThwError error, string message) : base(errorMessage[error] + " [" + message + "]")
        {
            Error = error;
        }
    }

    // ****************************************************************************************************
    /// <summary>
    /// Thw例外種別列挙子
    /// </summary>
    // ****************************************************************************************************
    public enum ThwError
    {
        THWNotFound,
        EditWindowNotFound,
        InvalidWindowHandle,
        InvalidArgument,
        NotSupportedFunction,
        FunctionAlreadyRunning,
        RequestTimeout,
    }

}
