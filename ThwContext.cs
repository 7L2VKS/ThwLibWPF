using Prism.Mvvm;
using System;

namespace ThwLib
{
    //
    // Model
    //
    public class ThwContext : BindableBase
    {
        private readonly Thw _thw;
        private IntPtr _hWnd;

        //
        // Constructor
        //
        public ThwContext()
        {
            _thw = Thw.GetInstance();
            _thw.SaveEvent += SaveResult;
        }

        //
        // Properties (for operations)
        //
        private Thw.WindowType _targetWindow;
        public Thw.WindowType TargetWindow
        {
            get => _targetWindow;
            set
            {
                SetProperty(ref _targetWindow, value);
                _thw.TargetWindow = _targetWindow;
            }
        }

        private bool _saving = false;
        public bool Saving
        {
            get => _saving;
            set => SetProperty(ref _saving, value);
        }

        private int _giveUpTime;
        public int GiveUpTime
        {
            get => _giveUpTime;
            set
            {
                SetProperty(ref _giveUpTime, value);
                _thw.GiveupTime = (int)_giveUpTime;
            }
        }

        private Exception? _thrownException = null;
        public Exception? ThrownException
        {
            get => _thrownException;
            set => SetProperty(ref _thrownException, value);
        }

        //
        // Properties (for setting data)
        //
        private string _setCall = string.Empty;
        public string SetCall
        {
            get => _setCall;
            set => SetProperty(ref _setCall, value);
        }

        private bool _setDX;
        public bool SetDX
        {
            get => _setDX;
            set => SetProperty(ref _setDX, value);
        }

        public DateTime _setDateTime;
        public DateTime SetDateTime
        {
            get => _setDateTime;
            set => SetProperty(ref _setDateTime, value);
        }

        private string _setHis = string.Empty;
        public string SetHis
        {
            get => _setHis;
            set => SetProperty(ref _setHis, value);
        }

        private string _setMy = string.Empty;
        public string SetMy
        {
            get => _setMy;
            set => SetProperty(ref _setMy, value);
        }

        private string _setFreq = string.Empty;
        public string SetFreq
        {
            get => _setFreq;
            set => SetProperty(ref _setFreq, value);
        }

        private string _setMode = string.Empty;
        public string SetMode
        {
            get => _setMode;
            set => SetProperty(ref _setMode, value);
        }

        private string _setCode = string.Empty;
        public string SetCode
        {
            get => _setCode;
            set => SetProperty(ref _setCode, value);
        }

        private string _setGL = string.Empty;
        public string SetGL
        {
            get => _setGL;
            set => SetProperty(ref _setGL, value);
        }

        private char? _setJ;
        public char? SetJ
        {
            get => _setJ;
            set => SetProperty(ref _setJ, value);
        }

        private char? _setS;
        public char? SetS
        {
            get => _setS;
            set => SetProperty(ref _setS, value);
        }

        private char? _setR;
        public char? SetR
        {
            get => _setR;
            set => SetProperty(ref _setR, value);
        }

        private string _setHisName = string.Empty;
        public string SetHisName
        {
            get => _setHisName;
            set => SetProperty(ref _setHisName, value);
        }

        private string _setQTH = string.Empty;
        public string SetQTH
        {
            get => _setQTH;
            set => SetProperty(ref _setQTH, value);
        }

        private bool _autoDXEntity;
        public bool AutoDXEntity
        {
            get => _autoDXEntity;
            set => SetProperty(ref _autoDXEntity, value);
        }

        private string _setRemarks1 = string.Empty;
        public string SetRemarks1
        {
            get => _setRemarks1;
            set => SetProperty(ref _setRemarks1, value);
        }

        private string _setRemarks2 = string.Empty;
        public string SetRemarks2
        {
            get => _setRemarks2;
            set => SetProperty(ref _setRemarks2, value);
        }

        private bool _setCQ;
        public bool SetCQ
        {
            get => _setCQ;
            set => SetProperty(ref _setCQ, value);
        }

        private bool _set1;
        public bool Set1
        {
            get => _set1;
            set => SetProperty(ref _set1, value);
        }

        private bool _set2;
        public bool Set2
        {
            get => _set2;
            set => SetProperty(ref _set2, value);
        }

        private bool _focus;
        public bool Focus
        {
            get => _focus;
            set => SetProperty(ref _focus, value);
        }

        private bool _enter;
        public bool Enter
        {
            get => _enter;
            set => SetProperty(ref _enter, value);
        }

        //
        // Properties (for getting data)
        //
        private string? _getCall;
        public string? GetCall
        {
            get => _getCall;
            set => SetProperty(ref _getCall, value);
        }

        private bool _getDX;
        public bool GetDX
        {
            get => _getDX;
            set => SetProperty(ref _getDX, value);
        }

        public string? _getDate;
        public string? GetDate
        {
            get => _getDate;
            set => SetProperty(ref _getDate, value);
        }

        public string? _getTime;
        public string? GetTime
        {
            get => _getTime;
            set => SetProperty(ref _getTime, value);
        }

        private string? _getHis;
        public string? GetHis
        {
            get => _getHis;
            set => SetProperty(ref _getHis, value);
        }

        private string? _getMy;
        public string? GetMy
        {
            get => _getMy;
            set => SetProperty(ref _getMy, value);
        }

        private string? _getFreq;
        public string? GetFreq
        {
            get => _getFreq;
            set => SetProperty(ref _getFreq, value);
        }

        private string? _getMode;
        public string? GetMode
        {
            get => _getMode;
            set => SetProperty(ref _getMode, value);
        }

        private string? _getCode;
        public string? GetCode
        {
            get => _getCode;
            set => SetProperty(ref _getCode, value);
        }

        private string? _getGL;
        public string? GetGL
        {
            get => _getGL;
            set => SetProperty(ref _getGL, value);
        }

        private char? _getJ;
        public char? GetJ
        {
            get => _getJ;
            set => SetProperty(ref _getJ, value);
        }

        private char? _getS;
        public char? GetS
        {
            get => _getS;
            set => SetProperty(ref _getS, value);
        }

        private char? _getR;
        public char? GetR
        {
            get => _getR;
            set => SetProperty(ref _getR, value);
        }

        private string? _getHisName;
        public string? GetHisName
        {
            get => _getHisName;
            set => SetProperty(ref _getHisName, value);
        }

        private string? _getQTH;
        public string? GetQTH
        {
            get => _getQTH;
            set => SetProperty(ref _getQTH, value);
        }

        private string? _getRemarks1;
        public string? GetRemarks1
        {
            get => _getRemarks1;
            set => SetProperty(ref _getRemarks1, value);
        }

        private string? _getRemarks2;
        public string? GetRemarks2
        {
            get => _getRemarks2;
            set => SetProperty(ref _getRemarks2, value);
        }

        private bool _getCQ;
        public bool GetCQ
        {
            get => _getCQ;
            set => SetProperty(ref _getCQ, value);
        }

        private bool _get1;
        public bool Get1
        {
            get => _get1;
            set => SetProperty(ref _get1, value);
        }

        private bool _get2;
        public bool Get2
        {
            get => _get2;
            set => SetProperty(ref _get2, value);
        }

        private HDBInfo? _getHDBInfo;
        public HDBInfo? GetHDBInfo
        {
            get => _getHDBInfo;
            set => SetProperty(ref _getHDBInfo, value);
        }

        private DirectionInfo? _getDirectionInfo;
        public DirectionInfo? GetDirectionInfo
        {
            get => _getDirectionInfo;
            set => SetProperty(ref _getDirectionInfo, value);
        }

        //
        // Command Execution (for operations)
        //
        public void ShowRecOnEdit(int recNumOnEdit) => _hWnd = _thw.ShowRecOnEdit(recNumOnEdit);
        public void Clear() => _hWnd = _thw.Clear();

        // データ登録（Saveボタン押下）処理（同期）
        // 同期的に実行する場合は、以下のコードを使用する。登録時にTHW上で確認ダイアログが表示される場合、それに応答するまで
        // UIスレッドに制御が戻らずロックされるため、非同期で実行することを推奨する。
        //public void Save(SaveConfirmationDialog op)
        //{
        //    if (TargetWindow == Thw.WindowType.Edit || op == SaveConfirmationDialog.Setting)
        //    {
        //        _hWnd = _thw.Save();
        //    }
        //    else
        //    {
        //        _hWnd = _thw.Save(op == SaveConfirmationDialog.Show);
        //    }   
        //}

        // データ登録（Saveボタン押下）処理（非同期）
        public async void Save(SaveConfirmationDialog op)
        {
            Saving = true;

            try
            {
                if (TargetWindow == Thw.WindowType.Edit || op == SaveConfirmationDialog.Setting)
                {
                    _hWnd = await _thw.SaveAsync();
                }
                else
                {
                    _hWnd = await _thw.SaveAsync(op == SaveConfirmationDialog.Show);
                }
            }
            catch (Exception ex)
            {
                ThrownException = ex;
            }
            finally
            {
                Saving = _thw.RunningSave;
            }
        }

        // データ登録（Saveボタン押下）処理（非同期）後、Give Upタイマーがタイムアップした後にTHWから制御が戻った場合に
        // 呼び出されるイベントハンドラ
        private void SaveResult(object sender, ThwSaveEventArgs e)
        {
            Saving = false;

            _hWnd = e.HWnd;
            ThrownException = e.Exception;
        }

        public void DupCheck() => _hWnd = _thw.DupCheck();
        public void Minimize() => _hWnd = _thw.Minimize();
        public void Restore() => _hWnd = _thw.Restore();

        //
        // Command Execution (for setting data)
        //
        public void DoSetCall() => _hWnd = _thw.SetCall(SetCall, Focus, Enter);
        public void DoSetDX() => _hWnd = _thw.SetDX(SetDX);
        public void DoSetDateTime() => _hWnd = _thw.SetDateTime(SetDateTime, Focus, Enter);
        public void DoSetDate() => _hWnd = _thw.SetDate(SetDateTime, Focus, Enter);
        public void DoSetTime() => _hWnd = _thw.SetTime(SetDateTime, Focus, Enter);
        public void DoSetHis() => _hWnd = _thw.SetHis(SetHis, Focus, Enter);
        public void DoSetMy() => _hWnd = _thw.SetMy(SetMy, Focus, Enter);
        public void DoSetFreq() => _hWnd = _thw.SetFreq(SetFreq, Focus, Enter);
        public void DoSetMode() => _hWnd = _thw.SetMode(SetMode, Focus, Enter);
        public void DoSetCode() => _hWnd = _thw.SetCode(SetCode, Focus, Enter);
        public void DoSetGL() => _hWnd = _thw.SetGL(SetGL, Focus, Enter);
        public void DoSetQSL()
        {
            string j = SetJ?.ToString() ?? " ";
            string s = SetS?.ToString() ?? " ";
            string r = SetR?.ToString() ?? " ";
            _hWnd = _thw.SetQSL((j + s + r).TrimEnd(), Focus, Enter);
        }
        public void DoSetJ() => _hWnd = _thw.SetQSL(SetJ, 0, Focus, Enter);
        public void DoSetS() => _hWnd = _thw.SetQSL(SetS, 1, Focus, Enter);
        public void DoSetR() => _hWnd = _thw.SetQSL(SetR, 2, Focus, Enter);
        public void DoSetHisName() => _hWnd = _thw.SetHisName(SetHisName, Focus, Enter);
        public void DoSetQTH() => _hWnd = _thw.SetQTH(SetQTH, Focus, Enter);
        public void DoSetAutoDXEntity() => _hWnd = _thw.SetAutoDXEntity(AutoDXEntity);
        public void DoSetRemarks1() => _hWnd = _thw.SetRemarks1(SetRemarks1, Focus, Enter);
        public void DoSetRemarks2() => _hWnd = _thw.SetRemarks2(SetRemarks2, Focus, Enter);
        public void DoSetCQ() => _hWnd = _thw.SetCQ(SetCQ);
        public void DoSet1() => _hWnd = _thw.Set1(Set1);
        public void DoSet2() => _hWnd = _thw.Set2(Set2);
        public void DoSetAll()
        {
            string j = SetJ?.ToString() ?? " ";
            string s = SetS?.ToString() ?? " ";
            string r = SetR?.ToString() ?? " ";
            var qso = new ThwQSORecord()
            {
                Call = SetCall,
                DateTime = SetDateTime,
                His = SetHis,
                My = SetMy,
                Freq = SetFreq,
                Mode = SetMode,
                Code = SetCode,
                GL = SetGL,
                QSL = (j + s + r).TrimEnd(),
                HisName = SetHisName,
                QTH = SetQTH,
                Remarks1 = SetRemarks1,
                Remarks2 = SetRemarks2,
                CheckDX = SetDX,
                CheckCQ = SetCQ,
                Check1 = Set1,
                Check2 = Set2
            };
            _hWnd = _thw.SetAll(qso, Focus, Enter);
        }

        //
        // Command Execution (for getting data)
        //
        public void DoGetCall()
        {
            var qso = new ThwQSORecord();
            _hWnd = _thw.GetAll(qso);
            GetCall = qso.Call;
            GetDX = qso.CheckDX;
        }

        public void DoGetDateTime()
        {
            _hWnd = _thw.GetDateTime(out string date, out string time);
            GetDate = date;
            GetTime = time;
        }

        public void DoGetDateTimeAsJST()
        {
            _hWnd = _thw.GetDateTime(out DateTime? dateTime, Thw.TimeZone.JST);
            ExtractDateTimeToProperties(dateTime);
        }

        public void DoGetDateTimeAsUTC()
        {
            _hWnd = _thw.GetDateTime(out DateTime? dateTime, Thw.TimeZone.UTC);
            ExtractDateTimeToProperties(dateTime);
        }

        private void ExtractDateTimeToProperties(DateTime? dateTime)
        {
            GetDate = ThwQSORecord.DateToString(dateTime);
            GetTime = ThwQSORecord.TimeToString(dateTime);
        }

        public void DoGetDate()
        {
            _hWnd = _thw.GetDate(out string str);
            GetDate = str;
        }

        public void DoGetTime()
        {
            _hWnd = _thw.GetTime(out string str);
            GetTime = str;
        }

        public void DoGetHis()
        {
            _hWnd = _thw.GetHis(out string str);
            GetHis = str;
        }

        public void DoGetMy()
        {
            _hWnd = _thw.GetMy(out string str);
            GetMy = str;
        }

        public void DoGetFreq()
        {
            _hWnd = _thw.GetFreq(out string str);
            GetFreq = str;
        }

        public void DoGetMode()
        {
            _hWnd = _thw.GetMode(out string str);
            GetMode = str;
        }

        public void DoGetCode()
        {
            _hWnd = _thw.GetCode(out string str);
            GetCode = str;
        }

        public void DoGetGL()
        {
            _hWnd = _thw.GetGL(out string str);
            GetGL = str;
        }

        public void DoGetQSL()
        {
            _hWnd = _thw.GetQSL(out string str);
            GetJ = str.Length >= 1 ? str[0] : null;
            GetS = str.Length >= 2 ? str[1] : null;
            GetR = str.Length >= 3 ? str[2] : null;
        }

        public void DoGetJ()
        {
            _hWnd = _thw.GetQSL(out char? ch, 0);
            GetJ = ch;
        }

        public void DoGetS()
        {
            _hWnd = _thw.GetQSL(out char? ch, 1);
            GetS = ch;
        }

        public void DoGetR()
        {
            _hWnd = _thw.GetQSL(out char? ch, 2);
            GetR = ch;
        }

        public void DoGetHisName()
        {
            _hWnd = _thw.GetHisName(out string str);
            GetHisName = str;
        }

        public void DoGetQTH()
        {
            _hWnd = _thw.GetQTH(out string str);
            GetQTH = str;
        }

        public void DoGetRemarks1()
        {
            _hWnd = _thw.GetRemarks1(out string str);
            GetRemarks1 = str;
        }

        public void DoGetRemarks2()
        {
            _hWnd = _thw.GetRemarks2(out string str);
            GetRemarks2 = str;
        }
        public void DoGetCQ12()
        {
            var qso = new ThwQSORecord();
            _hWnd = _thw.GetAll(qso);

            GetCQ = qso.CheckCQ;
            Get1 = qso.Check1;
            Get2 = qso.Check2;
        }

        public void DoGetAll()
        {
            var qso = new ThwQSORecord();
            _hWnd = _thw.GetAll(qso);
            ExtractQSOToProperties(qso);
        }

        public void DoGetAllAt()
        {
            var qso = new ThwQSORecord();
            _hWnd = _thw.GetAllAt(qso);
            ExtractQSOToProperties(qso);
        }

        public void DoGetAllAt(int recNum)
        {
            var qso = new ThwQSORecord();
            _hWnd = _thw.GetAllAt(qso, recNum);
            ExtractQSOToProperties(qso);
        }

        private void ExtractQSOToProperties(ThwQSORecord qso)
        {
            GetCall = qso.Call;
            GetDX = qso.CheckDX;
            GetDate = ThwQSORecord.DateToString(qso.DateTime);
            GetTime = ThwQSORecord.TimeToString(qso.DateTime);
            GetHis = qso.His;
            GetMy = qso.My;
            GetFreq = qso.Freq;
            GetMode = qso.Mode;
            GetCode = qso.Code;
            GetGL = qso.GL;
            GetJ = qso.QSL.Length >= 1 ? qso.QSL[0] : null;
            GetS = qso.QSL.Length >= 2 ? qso.QSL[1] : null;
            GetR = qso.QSL.Length >= 3 ? qso.QSL[2] : null;
            GetHisName = qso.HisName;
            GetQTH = qso.QTH;
            GetRemarks1 = qso.Remarks1;
            GetRemarks2 = qso.Remarks2;
            GetCQ = qso.CheckCQ;
            Get1 = qso.Check1;
            Get2 = qso.Check2;
        }

        public void DoGetHDBInfo()
        {
            _hWnd = _thw.GetHDBInfo(out string path, out int count);
            GetHDBInfo = new HDBInfo() { Path = path, Count = count };
        }

        public void DoGetDirectionInfo()
        {
            _hWnd = _thw.GetDirction(out int? direction, out float? distance);
            GetDirectionInfo = new DirectionInfo() { Direction = direction, Distance = distance };
        }

        public class HDBInfo
        {
            public required string Path { get; init; }
            public int Count { get; init; }
        }

        public class DirectionInfo
        {
            public int? Direction { get; init; }
            public float? Distance { get; init; }
        }

        public enum SaveConfirmationDialog
        {
            Setting,
            Show,
            NoShow
        }
    }
}
