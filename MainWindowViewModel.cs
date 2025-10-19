using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace ThwLib
{
    //
    // ViewModel
    //
    public class MainWindowViewModel : BindableBase
    {
        private readonly ThwContext _thwContext;   // Model
        private readonly IDialogService _dialogService;

        //
        // Constructor
        //
        public MainWindowViewModel(IDialogService dialogService)
        {
            _thwContext = new ThwContext();
            _thwContext.PropertyChanged += ThwContextPropertyChanged;
            _dialogService = dialogService;

            ExecuteSetCommand = new DelegateCommand<String>(ExecuteSet);
            ExecuteGetCommand = new DelegateCommand<String>(ExecuteGet);
            ExecuteOpeCommand = new DelegateCommand<String>(ExecuteOpe, CanExecuteOpe);
        }

        //
        // Commands
        //
        public DelegateCommand<string> ExecuteSetCommand { get; private set; }
        public DelegateCommand<string> ExecuteGetCommand { get; private set; }
        public DelegateCommand<string> ExecuteOpeCommand { get; private set; }

        //
        // Properties (for operations)
        //
        private Thw.WindowType _targetWindow = Thw.WindowType.Input;
        public Thw.WindowType TargetWindow
        {
            get => _targetWindow;
            set => SetProperty(ref _targetWindow, value);
        }

        private decimal _recNumOnEdit = 1;
        public decimal RecNumOnEdit
        {
            get => _recNumOnEdit;
            set => SetProperty(ref _recNumOnEdit, value);
        }

        private ThwContext.SaveConfirmationDialog _saveOption = ThwContext.SaveConfirmationDialog.Setting;
        public ThwContext.SaveConfirmationDialog SaveOption
        {
            get => _saveOption;
            set => SetProperty(ref _saveOption, value);
        }

        private bool _canSave = true;
        public bool CanSave
        {
            get => _canSave;
            set
            {
                SetProperty(ref _canSave, value);
                ExecuteOpeCommand.RaiseCanExecuteChanged();
            }
        }

        private decimal _giveUpTime = 0;
        public decimal GiveUpTime
        {
            get => _giveUpTime;
            set => SetProperty(ref _giveUpTime, value);
        }

        //
        // Properties (for setting data)
        //
        private string _setCall = string.Empty;
        public string SetCall
        {
            get => _setCall;
            set
            {
                SetProperty(ref _setCall, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private bool _setDX = false;
        public bool SetDX
        {
            get => _setDX;
            set => SetProperty(ref _setDX, value);
        }

        public DateTime _setDateTime = DateTime.Now;

        public DateTime SetDateTime
        {
            get => _setDateTime;
            set
            {
                SetProperty(ref _setDateTime, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private bool _setUTC = false;
        public bool SetUTC
        {
            get => _setUTC;
            set => SetProperty(ref _setUTC, value);
        }

        private string _setHis = string.Empty;
        public string SetHis
        {
            get => _setHis;
            set
            {
                SetProperty(ref _setHis, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setMy = string.Empty;
        public string SetMy
        {
            get => _setMy;
            set
            {
                SetProperty(ref _setMy, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setFreq = string.Empty;
        public string SetFreq
        {
            get => _setFreq;
            set
            {
                SetProperty(ref _setFreq, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setMode = string.Empty;
        public string SetMode
        {
            get => _setMode;
            set
            {
                SetProperty(ref _setMode, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setCode = string.Empty;
        public string SetCode
        {
            get => _setCode;
            set
            {
                SetProperty(ref _setCode, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setGL = string.Empty;
        public string SetGL
        {
            get => _setGL;
            set
            {
                SetProperty(ref _setGL, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setJ = string.Empty;
        public string SetJ
        {
            get => _setJ;
            set
            {
                SetProperty(ref _setJ, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setS = string.Empty;
        public string SetS
        {
            get => _setS;
            set
            {
                SetProperty(ref _setS, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setR = string.Empty;
        public string SetR
        {
            get => _setR;
            set
            {
                SetProperty(ref _setR, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setHisName = string.Empty;
        public string SetHisName
        {
            get => _setHisName;
            set
            {
                SetProperty(ref _setHisName, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setQTH = string.Empty;
        public string SetQTH
        {
            get => _setQTH;
            set
            {
                SetProperty(ref _setQTH, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private bool _autoDXEntity = true;
        public bool AutoDXEntity
        {
            get => _autoDXEntity;
            set
            {
                SetProperty(ref _autoDXEntity, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setRemarks1 = string.Empty;
        public string SetRemarks1
        {
            get => _setRemarks1;
            set
            {
                SetProperty(ref _setRemarks1, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private string _setRemarks2 = string.Empty;
        public string SetRemarks2
        {
            get => _setRemarks2;
            set
            {
                SetProperty(ref _setRemarks2, value);
                ExecuteSetCommand.RaiseCanExecuteChanged();
            }
        }

        private bool _setCQ = false;
        public bool SetCQ
        {
            get => _setCQ;
            set => SetProperty(ref _setCQ, value);
        }

        private bool _set1 = false;
        public bool Set1
        {
            get => _set1;
            set => SetProperty(ref _set1, value);
        }

        private bool _set2 = false;
        public bool Set2
        {
            get => _set2;
            set => SetProperty(ref _set2, value);
        }

        private bool _focus = false;
        public bool Focus
        {
            get => _focus;
            set => SetProperty(ref _focus, value);
        }

        private bool _enter = false;
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

        private TimeZoneSelection _selectedTimeZone = TimeZoneSelection.AsIs;
        public TimeZoneSelection SelectedTimeZone
        {
            get => _selectedTimeZone;
            set => SetProperty(ref _selectedTimeZone, value);
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

        private string? _getJ;
        public string? GetJ
        {
            get => _getJ;
            set => SetProperty(ref _getJ, value);
        }

        private string? _getS;
        public string? GetS
        {
            get => _getS;
            set => SetProperty(ref _getS, value);
        }

        private string? _getR;
        public string? GetR
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

        private RecordSelection _selectedRecord = RecordSelection.InputEdit;
        public RecordSelection SelectedRecord
        {
            get => _selectedRecord;
            set => SetProperty(ref _selectedRecord, value);
        }

        private decimal _recNumToGet = 1;
        public decimal RecNumToGet
        {
            get => _recNumToGet;
            set => SetProperty(ref _recNumToGet, value);
        }

        public ObservableCollection<ComboBoxItem> InfoItems { get; } = [
            new ("HAMLOG.HDBファイル情報", InfoSelection.HDB),
            new ("方位・距離表示", InfoSelection.Direction)
        ];

        private ComboBoxItem? _selectedInfoItem;
        public ComboBoxItem? SelectedInfoItem
        {
            get => _selectedInfoItem;
            set => SetProperty(ref _selectedInfoItem, value);
        }

        private ThwContext.HDBInfo? _getHDBInfo;
        public ThwContext.HDBInfo? GetHDBInfo
        {
            get => _getHDBInfo;
            set
            {
                SetProperty(ref _getHDBInfo, value);
                if (_getHDBInfo != null)
                {
                    string msg = $"File path = {_getHDBInfo.Path}{Environment.NewLine}Record count = {_getHDBInfo.Count}";
                    _dialogService.ShowMessage(msg, IDialogService.Severity.Information);
                }
            }
        }

        private ThwContext.DirectionInfo? _getDirectionInfo;
        public ThwContext.DirectionInfo? GetDirectionInfo
        {
            get => _getDirectionInfo;
            set
            {
                SetProperty(ref _getDirectionInfo, value);
                if (_getDirectionInfo != null)
                {
                    string strDirection = _getDirectionInfo.Direction?.ToString() ?? "Unknown";
                    string strDistance = _getDirectionInfo.Distance?.ToString() ?? "Unknown";
                    string msg = $"Direction = {strDirection}{Environment.NewLine}Distance = {strDistance}";
                    _dialogService.ShowMessage(msg, IDialogService.Severity.Information);
                }
            }
        }

        //
        // Binding with Model
        //
        private void ThwContextPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "Saving":
                    CanSave = !_thwContext.Saving;
                    break;
                case "ThrownException":
                    Exception? ex = _thwContext.ThrownException;
                    if (ex != null)
                    {
                        _dialogService.ShowMessage($"{ex.GetType()} : {ex.Message}", IDialogService.Severity.Error);
                    }
                    break;
                case "GetCall":
                    GetCall = _thwContext.GetCall;
                    break;
                case "GetDX":
                    GetDX = _thwContext.GetDX;
                    break;
                case "GetDate":
                    GetDate = _thwContext.GetDate;
                    break;
                case "GetTime":
                    GetTime = _thwContext.GetTime;
                    break;
                case "GetHis":
                    GetHis = _thwContext.GetHis;
                    break;
                case "GetMy":
                    GetMy = _thwContext.GetMy;
                    break;
                case "GetFreq":
                    GetFreq = _thwContext.GetFreq;
                    break;
                case "GetMode":
                    GetMode = _thwContext.GetMode;
                    break;
                case "GetCode":
                    GetCode = _thwContext.GetCode;
                    break;
                case "GetGL":
                    GetGL = _thwContext.GetGL;
                    break;
                case "GetJ":
                    GetJ = _thwContext.GetJ?.ToString() ?? string.Empty;
                    break;
                case "GetS":
                    GetS = _thwContext.GetS?.ToString() ?? string.Empty;
                    break;
                case "GetR":
                    GetR = _thwContext.GetR?.ToString() ?? string.Empty;
                    break;
                case "GetHisName":
                    GetHisName = _thwContext.GetHisName;
                    break;
                case "GetQTH":
                    GetQTH = _thwContext.GetQTH;
                    break;
                case "GetRemarks1":
                    GetRemarks1 = _thwContext.GetRemarks1;
                    break;
                case "GetRemarks2":
                    GetRemarks2 = _thwContext.GetRemarks2;
                    break;
                case "GetCQ":
                    GetCQ = _thwContext.GetCQ;
                    break;
                case "Get1":
                    Get1 = _thwContext.Get1;
                    break;
                case "Get2":
                    Get2 = _thwContext.Get2;
                    break;
                case "GetHDBInfo":
                    GetHDBInfo = _thwContext.GetHDBInfo;
                    break;
                case "GetDirectionInfo":
                    GetDirectionInfo = _thwContext.GetDirectionInfo;
                    break;
            }
        }

        //
        // Command Execution (for operations)
        //
        private void ExecuteOpe(string command)
        {
            _thwContext.TargetWindow = TargetWindow;

            try
            {
                switch (command)
                {
                    case "ShowRecOnEdit":
                        _thwContext.ShowRecOnEdit((int)RecNumOnEdit);
                        break;
                    case "Clear":
                        _thwContext.Clear();
                        break;
                    case "Save":
                        _thwContext.GiveUpTime = (int)GiveUpTime;
                        _thwContext.Save(SaveOption);
                        break;
                    case "DupCheck":
                        _thwContext.DupCheck();
                        break;
                    case "Minimize":
                        _thwContext.Minimize();
                        break;
                    case "Restore":
                        _thwContext.Restore();
                        break;
                }
            }
            catch (Exception ex)
            {
                _dialogService.ShowMessage($"{ex.GetType()} : {ex.Message}", IDialogService.Severity.Error);
            }
        }
        
        private bool CanExecuteOpe(string command)
        {
            return command switch
            {
                "Save" => CanSave,
                _ => true,
            };
        }

        //
        // Command Execution (for setting data)
        //
        private void ExecuteSet(string command)
        {
            _thwContext.TargetWindow = TargetWindow;
            _thwContext.Focus = Focus;
            _thwContext.Enter = Enter;

            try
            {
                switch (command)
                {
                    case "Call":
                        _thwContext.SetCall = SetCall;
                        _thwContext.DoSetCall();
                        break;
                    case "DX":
                        _thwContext.SetDX = SetDX;
                        _thwContext.DoSetDX();
                        break;
                    case "DateTime":
                        _thwContext.SetDateTime = SetDateTime;
                        _thwContext.DoSetDateTime();
                        break;
                    case "Date":
                        _thwContext.SetDateTime = SetDateTime;
                        _thwContext.DoSetDate();
                        break;
                    case "Time":
                        _thwContext.SetDateTime = SetDateTime;
                        _thwContext.DoSetTime();
                        break;
                    case "UTC":
                        SetDateTime = SetUTC ? SetDateTime.ToUniversalTime() : SetDateTime.ToLocalTime();
                        _thwContext.SetDateTime = SetDateTime;
                        break;
                    case "His":
                        _thwContext.SetHis = SetHis;
                        _thwContext.DoSetHis();
                        break;
                    case "My":
                        _thwContext.SetMy = SetMy;
                        _thwContext.DoSetMy();
                        break;
                    case "Freq":
                        _thwContext.SetFreq = SetFreq;
                        _thwContext.DoSetFreq();
                        break;
                    case "Mode":
                        _thwContext.SetMode = SetMode;
                        _thwContext.DoSetMode();
                        break;
                    case "Code":
                        _thwContext.SetCode = SetCode;
                        _thwContext.DoSetCode();
                        break;
                    case "GL":
                        _thwContext.SetGL = SetGL;
                        _thwContext.DoSetGL();
                        break;
                    case "QSL":
                        _thwContext.SetJ = string.IsNullOrEmpty(SetJ) ? null : SetJ[0];
                        _thwContext.SetS = string.IsNullOrEmpty(SetS) ? null : SetS[0];
                        _thwContext.SetR = string.IsNullOrEmpty(SetR) ? null : SetR[0];
                        _thwContext.DoSetQSL();
                        break;
                    case "J":
                        _thwContext.SetJ = string.IsNullOrEmpty(SetJ) ? null : SetJ[0];
                        _thwContext.DoSetJ();
                        break;
                    case "S":
                        _thwContext.SetS = string.IsNullOrEmpty(SetS) ? null : SetS[0];
                        _thwContext.DoSetS();
                        break;
                    case "R":
                        _thwContext.SetR = string.IsNullOrEmpty(SetR) ? null : SetR[0];
                        _thwContext.DoSetR();
                        break;
                    case "HisName":
                        _thwContext.SetHisName = SetHisName;
                        _thwContext.DoSetHisName();
                        break;
                    case "QTH":
                        _thwContext.SetQTH = SetQTH;
                        _thwContext.DoSetQTH();
                        break;
                    case "AutoDXEntity":
                        _thwContext.AutoDXEntity = AutoDXEntity;
                        _thwContext.DoSetAutoDXEntity();
                        break;
                    case "Remarks1":
                        _thwContext.SetRemarks1 = SetRemarks1;
                        _thwContext.DoSetRemarks1();
                        break;
                    case "Remarks2":
                        _thwContext.SetRemarks2 = SetRemarks2;
                        _thwContext.DoSetRemarks2();
                        break;
                    case "CQ":
                        _thwContext.SetCQ = SetCQ;
                        _thwContext.DoSetCQ();
                        break;
                    case "1":
                        _thwContext.Set1 = Set1;
                        _thwContext.DoSet1();
                        break;
                    case "2":
                        _thwContext.Set2 = Set2;
                        _thwContext.DoSet2();
                        break;
                    case "All":
                        _thwContext.SetCall = SetCall;
                        _thwContext.SetDX = SetDX;
                        _thwContext.SetDateTime = SetDateTime;
                        _thwContext.SetHis = SetHis;
                        _thwContext.SetMy = SetMy;
                        _thwContext.SetFreq = SetFreq;
                        _thwContext.SetMode = SetMode;
                        _thwContext.SetCode = SetCode;
                        _thwContext.SetGL = SetGL;
                        _thwContext.SetJ = string.IsNullOrEmpty(SetJ) ? null : SetJ[0];
                        _thwContext.SetS = string.IsNullOrEmpty(SetS) ? null : SetS[0];
                        _thwContext.SetR = string.IsNullOrEmpty(SetR) ? null : SetR[0];
                        _thwContext.SetHisName = SetHisName;
                        _thwContext.SetQTH = SetQTH;
                        _thwContext.SetRemarks1 = SetRemarks1;
                        _thwContext.SetRemarks2 = SetRemarks2;
                        _thwContext.SetCQ = SetCQ;
                        _thwContext.Set1 = Set1;
                        _thwContext.Set2 = Set2;
                        _thwContext.DoSetAll();
                        break;
                }
            }
            catch (Exception ex)
            {
                _dialogService.ShowMessage($"{ex.GetType()} : {ex.Message}", IDialogService.Severity.Error);
            }
        }

        //
        // Command Execution (for getting data)
        //
        private void ExecuteGet(string command)
        {
            _thwContext.TargetWindow = TargetWindow;

            try
            {
                switch (command)
                {
                    case "Call":
                        _thwContext.DoGetCall();
                        break;
                    case "DateTime":
                        switch (SelectedTimeZone)
                        {
                            case TimeZoneSelection.AsIs:
                                _thwContext.DoGetDateTime();
                                break;
                            case TimeZoneSelection.AsJST:
                                _thwContext.DoGetDateTimeAsJST();
                                break;
                            case TimeZoneSelection.AsUTC:
                                _thwContext.DoGetDateTimeAsUTC();
                                break;
                        }
                        break;
                    case "Date":
                        _thwContext.DoGetDate();
                        break;
                    case "Time":
                        _thwContext.DoGetTime();
                        break;
                    case "His":
                        _thwContext.DoGetHis();
                        break;
                    case "My":
                        _thwContext.DoGetMy();
                        break;
                    case "Freq":
                        _thwContext.DoGetFreq();
                        break;
                    case "Mode":
                        _thwContext.DoGetMode();
                        break;
                    case "Code":
                        _thwContext.DoGetCode();
                        break;
                    case "GL":
                        _thwContext.DoGetGL();
                        break;
                    case "QSL":
                        _thwContext.DoGetQSL();
                        break;
                    case "J":
                        _thwContext.DoGetJ();
                        break;
                    case "S":
                        _thwContext.DoGetS();
                        break;
                    case "R":
                        _thwContext.DoGetR();
                        break;
                    case "HisName":
                        _thwContext.DoGetHisName();
                        break;
                    case "QTH":
                        _thwContext.DoGetQTH();
                        break;
                    case "Remarks1":
                        _thwContext.DoGetRemarks1();
                        break;
                    case "Remarks2":
                        _thwContext.DoGetRemarks2();
                        break;
                    case "CQ12":
                        _thwContext.DoGetCQ12();
                        break;
                    case "All":
                        switch (SelectedRecord)
                        {
                            case RecordSelection.InputEdit:
                                _thwContext.DoGetAll();
                                break;
                            case RecordSelection.Main:
                                _thwContext.DoGetAllAt();
                                break;
                            case RecordSelection.Number:
                                _thwContext.DoGetAllAt((int)RecNumToGet);
                                break;
                        }
                        break;
                    case "Info":
                        if (SelectedInfoItem != null)
                        {
                            switch ((InfoSelection)SelectedInfoItem.Index)
                            {
                                case InfoSelection.HDB:
                                    _thwContext.DoGetHDBInfo();
                                    break;
                                case InfoSelection.Direction:
                                    _thwContext.DoGetDirectionInfo();
                                    break;
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                _dialogService.ShowMessage($"{ex.GetType()} : {ex.Message}", IDialogService.Severity.Error);
            }
        }

        //
        // Enum for timezone designation
        //
        public enum TimeZoneSelection
        {
            AsIs,
            AsJST,
            AsUTC
        }

        //
        // Enum for record specifier to get
        //
        public enum RecordSelection
        {
            InputEdit,
            Main,
            Number
        }

        //
        // Enum for information to get
        //
        public enum InfoSelection
        {
            HDB,
            Direction
        }

        //
        // Class for ComboBox items
        //
        public class ComboBoxItem(string label, object index)
        {
            public string Label { get; private set; } = label;
            public object Index { get; private set; } = index;
        }
    }
}
