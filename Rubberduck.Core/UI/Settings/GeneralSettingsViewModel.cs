using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Settings;
using Rubberduck.Common;
using Rubberduck.Interaction;
using NLog;
using Rubberduck.Api;
using Rubberduck.Git;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.VbeRuntime.Settings;
using Rubberduck.Resources;
using Rubberduck.Resources.Settings;
using Rubberduck.Parsing.Common;
using Rubberduck.UI.WPF;

namespace Rubberduck.UI.Settings
{
    public enum DelimiterOptions
    {
        Period = 46,
        Slash = 47
    }

    public sealed class GeneralSettingsViewModel : SettingsViewModelBase<Rubberduck.Settings.GeneralSettings>, ISettingsViewModel<Rubberduck.Settings.GeneralSettings>
    {
        private readonly IOperatingSystem _operatingSystem;
        private readonly IMessageBox _messageBox;
        private readonly IVbeSettings _vbeSettings;
        private readonly IProcDumpInstaller _procDumpInstaller;
        private readonly IGitProvisioner _gitProvisioner;
        private readonly IConfigurationService<HotkeySettings> _hotkeyService;

        private bool _indenterPrompted;
        private readonly IReadOnlyList<Type> _experimentalFeatureTypes;

        public GeneralSettingsViewModel(
            Configuration config,
            IOperatingSystem operatingSystem,
            IMessageBox messageBox,
            IVbeSettings vbeSettings,
            IExperimentalTypesProvider experimentalTypesProvider,
            IConfigurationService<Rubberduck.Settings.GeneralSettings> service,
            IConfigurationService<HotkeySettings> hotkeyService)
            : this(config, operatingSystem, messageBox, vbeSettings, experimentalTypesProvider, service, hotkeyService, null)
        {
        }

        public GeneralSettingsViewModel(
            Configuration config,
            IOperatingSystem operatingSystem,
            IMessageBox messageBox,
            IVbeSettings vbeSettings,
            IExperimentalTypesProvider experimentalTypesProvider,
            IConfigurationService<Rubberduck.Settings.GeneralSettings> service,
            IConfigurationService<HotkeySettings> hotkeyService,
            ApiConnection apiConnection,
            IFirewallCheck firewallCheck = null,
            IProcDumpInstaller procDumpInstaller = null,
            IGitProvisioner gitProvisioner = null)
            : base(service)
        {
            _gitProvisioner = gitProvisioner;
            _firewallCheck = firewallCheck;
            _procDumpInstaller = procDumpInstaller;
            _apiConnection = apiConnection;
            _operatingSystem = operatingSystem;
            _messageBox = messageBox;
            _vbeSettings = vbeSettings;
            _experimentalFeatureTypes = experimentalTypesProvider.ExperimentalTypes;

            Languages = new ObservableCollection<DisplayLanguageSetting>(Locales.AvailableCultures
                .OrderBy(locale => locale.NativeName)
                .Select(locale => new DisplayLanguageSetting(locale.Name)));

            LogLevels = new ObservableCollection<MinimumLogLevel>(
                LogLevelHelper.LogLevels.Select(l => new MinimumLogLevel(l.Ordinal, l.Name)));
            TransferSettingsToView(config.UserSettings.GeneralSettings, config.UserSettings.HotkeySettings);

            ShowLogFolderCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ShowLogFolder());
            ExportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ExportSettings(GetCurrentGeneralSettings()));
            ImportButtonCommand = new DelegateCommand(LogManager.GetCurrentClassLogger(), _ => ImportSettings());

            _hotkeyService = hotkeyService;
        }

        public List<ExperimentalFeature> ExperimentalFeatures { get; set; }

        public ObservableCollection<DisplayLanguageSetting> Languages { get; set; } 

        private DisplayLanguageSetting _selectedLanguage;
        public DisplayLanguageSetting SelectedLanguage
        {
            get => _selectedLanguage;
            set
            {
                if (!Equals(_selectedLanguage, value))
                {
                    _selectedLanguage = value;
                    OnPropertyChanged();
                }
            }
        }

        private ObservableViewModelCollection<HotkeySettingViewModel> _hotkeys;
        public ObservableViewModelCollection<HotkeySettingViewModel> Hotkeys
        {
            get => _hotkeys;
            set
            {
                if (_hotkeys != value)
                {
                    if (_hotkeys != null)
                    {
                        _hotkeys.ElementPropertyChanged -= InvalidateShouldDisplayWarning;
                    }
                    if (value != null)
                    {
                        value.ElementPropertyChanged += InvalidateShouldDisplayWarning;
                    }
                    _hotkeys = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ShouldDisplayHotkeyModificationLabel));
                }
            }
        }

        private void InvalidateShouldDisplayWarning(object sender, ElementPropertyChangedEventArgs<HotkeySettingViewModel> e)
        {
            OnPropertyChanged(nameof(ShouldDisplayHotkeyModificationLabel));
        }

        public bool ShouldDisplayHotkeyModificationLabel
        {
            get
            {
                return _hotkeys.Any(s => !s.IsValid && s.IsEnabled);
            }
        }

        private bool _autoSaveEnabled;
        public bool AutoSaveEnabled
        {
            get => _autoSaveEnabled;
            set
            {
                if (_autoSaveEnabled != value)
                {
                    _autoSaveEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _showSplashAtStartup;
        public bool ShowSplashAtStartup
        {
            get => _showSplashAtStartup;
            set
            {
                if (_showSplashAtStartup != value)
                {
                    _showSplashAtStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _checkVersionAtStartup;
        public bool CheckVersionAtStartup
        {
            get => _checkVersionAtStartup;
            set
            {
                if (_checkVersionAtStartup != value)
                {
                    _checkVersionAtStartup = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _includePreRelease;
        public bool IncludePreRelease
        {
            get => _includePreRelease;
            set
            {
                if (_includePreRelease != value)
                {
                    _includePreRelease = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _compileBeforeParse;
        public bool CompileBeforeParse
        {
            get => _compileBeforeParse;
            set
            {
                if (_compileBeforeParse == value)
                {
                    return;
                }

                if (value && _vbeSettings.CompileOnDemand)
                {
                    if(!SynchronizeVbeSettings())
                    {
                        return;
                    }
                }

                _compileBeforeParse = value;
                OnPropertyChanged();
            }
        }

        private bool _enableFolderDragAndDrop;
        public bool EnableFolderDragAndDrop
        {
            get => _enableFolderDragAndDrop;
            set
            {
                if (_enableFolderDragAndDrop != value)
                {
                    _enableFolderDragAndDrop = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _apiEnabled;
        public bool ApiEnabled
        {
            get => _apiEnabled;
            set
            {
                if (_apiEnabled != value)
                {
                    _apiEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _apiAllowNetwork;
        public bool ApiAllowNetwork
        {
            get => _apiAllowNetwork;
            set
            {
                if (_apiAllowNetwork != value)
                {
                    _apiAllowNetwork = value;
                    OnPropertyChanged();
                    RefreshFirewallHint();
                }
            }
        }

        private bool _apiAllowImmediate;
        public bool ApiAllowImmediate
        {
            get => _apiAllowImmediate;
            set
            {
                if (_apiAllowImmediate != value)
                {
                    _apiAllowImmediate = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _apiAllowImmediateRead;
        public bool ApiAllowImmediateRead
        {
            get => _apiAllowImmediateRead;
            set
            {
                if (_apiAllowImmediateRead != value)
                {
                    _apiAllowImmediateRead = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _apiAllowProcDump;
        public bool ApiAllowProcDump
        {
            get => _apiAllowProcDump;
            set
            {
                if (_apiAllowProcDump != value)
                {
                    _apiAllowProcDump = value;
                    OnPropertyChanged();
                }
            }
        }

        private CommandBase _ensureProcDumpCommand;
        /// <summary>Run when the user ticks the ProcDump option: if ProcDump is not there, offers to download it, and takes the tick back if it is not to be had.</summary>
        public CommandBase EnsureProcDumpCommand => _ensureProcDumpCommand ?? (_ensureProcDumpCommand = new DelegateCommand(
            LogManager.GetCurrentClassLogger(), _ => EnsureProcDump(), _ => _procDumpInstaller != null));

        private void EnsureProcDump()
        {
            if (!ApiAllowProcDump || _procDumpInstaller.Find() != null)
            {
                return;
            }

            if (!_messageBox.ConfirmYesNo(string.Format(GeneralSettingsUI.ApiAllowProcDump_Download, ProcDumpInstaller.DownloadUrl, _procDumpInstaller.TargetFolder),
                GeneralSettingsUI.ApiAllowProcDump_DownloadCaption, true))
            {
                ApiAllowProcDump = false;
                return;
            }

            try
            {
                _procDumpInstaller.Install();
            }
            catch (System.IO.IOException exception)
            {
                LogManager.GetCurrentClassLogger().Warn(exception, "ProcDump could not be installed.");
                _messageBox.NotifyWarn(exception.Message, GeneralSettingsUI.ApiAllowProcDump_DownloadCaption);
                ApiAllowProcDump = false;
            }
        }

        private bool _gitClientEnabled;
        public bool GitClientEnabled
        {
            get => _gitClientEnabled;
            set
            {
                if (_gitClientEnabled != value)
                {
                    _gitClientEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _gitPath;
        public string GitPath
        {
            get => _gitPath;
            set
            {
                if (_gitPath != value)
                {
                    _gitPath = value;
                    OnPropertyChanged();
                }
            }
        }

        private bool _gitAiEnabled;
        public bool GitAiEnabled
        {
            get => _gitAiEnabled;
            set
            {
                if (_gitAiEnabled != value)
                {
                    _gitAiEnabled = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _gitAiModel;
        public string GitAiModel
        {
            get => _gitAiModel;
            set
            {
                if (_gitAiModel != value)
                {
                    _gitAiModel = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>Where the API key is kept. Not the settings file: the key is a secret.</summary>
        public Rubberduck.Git.ICredentialStore Credentials { get; set; } = new Rubberduck.Git.WindowsCredentialStore();

        private static readonly string AiTarget = Rubberduck.Git.CredentialTargets.Ai("anthropic");

        public string GitAiKeyStatus => Credentials.Read(AiTarget) != null ? GeneralSettingsUI.GitAi_KeySaved : GeneralSettingsUI.GitAi_KeyMissing;

        private CommandBase _setGitAiKeyCommand;
        public CommandBase SetGitAiKeyCommand => _setGitAiKeyCommand ?? (_setGitAiKeyCommand = new DelegateCommand(
            LogManager.GetCurrentClassLogger(), _ =>
            {
                var key = Rubberduck.UI.CodeExplorer.GitClient.GitAiKeyDialog.Ask();
                if (key != null)
                {
                    Credentials.Write(AiTarget, null, key);
                    OnPropertyChanged(nameof(GitAiKeyStatus));
                }
            }));

        private CommandBase _removeGitAiKeyCommand;
        public CommandBase RemoveGitAiKeyCommand => _removeGitAiKeyCommand ?? (_removeGitAiKeyCommand = new DelegateCommand(
            LogManager.GetCurrentClassLogger(), _ =>
            {
                Credentials.Delete(AiTarget);
                OnPropertyChanged(nameof(GitAiKeyStatus));
            }));

        private CommandBase _ensureGitCommand;
        /// <summary>Run when the user ticks the git client option: if there is no usable git, offers to download one, and takes the tick back if none is to be had.</summary>
        public CommandBase EnsureGitCommand => _ensureGitCommand ?? (_ensureGitCommand = new DelegateCommand(
            LogManager.GetCurrentClassLogger(), _ => EnsureGit(), _ => _gitProvisioner != null));

        private void EnsureGit()
        {
            if (!GitClientEnabled || _gitProvisioner.Find(GitPath) != null)
            {
                return;
            }

            var megabytes = Math.Round(_gitProvisioner.DownloadSize / 1048576d);
            if (!_messageBox.ConfirmYesNo(string.Format(GeneralSettingsUI.GitClient_Download, _gitProvisioner.DownloadUrl, megabytes, _gitProvisioner.TargetFolder),
                GeneralSettingsUI.GitClient_DownloadCaption, true))
            {
                GitClientEnabled = false;
                return;
            }

            try
            {
                _gitProvisioner.Install();
            }
            catch (System.IO.IOException exception)
            {
                LogManager.GetCurrentClassLogger().Warn(exception, "Git could not be installed.");
                _messageBox.NotifyWarn(exception.Message + Environment.NewLine + Environment.NewLine + GeneralSettingsUI.GitClient_DownloadFailed, GeneralSettingsUI.GitClient_DownloadCaption);
                GitClientEnabled = false;
            }
        }

        private int _apiPort;
        public int ApiPort
        {
            get => _apiPort;
            set
            {
                if (_apiPort != value)
                {
                    _apiPort = value;
                    OnPropertyChanged();
                    RefreshFirewallHint();
                }
            }
        }

        //Null where the API is not available, as in tests.
        private readonly IFirewallCheck _firewallCheck;

        private string _firewallHint;
        /// <summary>What Windows Firewall will do with agents on the network, shown beside the network option; empty if the option is off or nothing is known.</summary>
        public string FirewallHint
        {
            get => _firewallHint;
            private set
            {
                if (_firewallHint != value)
                {
                    _firewallHint = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(HasFirewallHint));
                }
            }
        }

        public bool HasFirewallHint => !string.IsNullOrEmpty(_firewallHint);

        private int _firewallHintVersion;
        private void RefreshFirewallHint()
        {
            var version = ++_firewallHintVersion;
            if (_firewallCheck == null || !_apiAllowNetwork)
            {
                FirewallHint = null;
                return;
            }

            var port = _apiPort > 0 && _apiPort <= 65535 ? _apiPort : ApiServer.DefaultPort;
            //The firewall is queried off the UI thread; a newer change makes this answer stale.
            System.Threading.Tasks.Task.Run(() => FirewallHintText(_firewallCheck.Check(port)))
                .ContinueWith(task =>
                {
                    if (version == _firewallHintVersion && !task.IsFaulted)
                    {
                        FirewallHint = task.Result;
                    }
                });
        }

        public static string FirewallHintText(FirewallStatus status)
        {
            if (status == null || status.Profile == FirewallCheck.Unknown && status.Firewall == FirewallCheck.Unknown)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(status.Hint))
            {
                return status.Hint;
            }

            return status.NetworkReachable == "yes"
                ? $"Windows Firewall ({status.Profile} network): other machines can reach port {status.Port}."
                : $"Windows Firewall ({status.Profile} network): could not tell whether port {status.Port} is open.";
        }

        //Null where the API is not available, as in tests.
        private readonly ApiConnection _apiConnection;

        private CommandBase _copyApiConnectionCommand;
        public CommandBase CopyApiConnectionCommand => _copyApiConnectionCommand ?? (_copyApiConnectionCommand = new DelegateCommand(
            LogManager.GetCurrentClassLogger(), _ => System.Windows.Clipboard.SetText(_apiConnection.Details()), _ => _apiConnection != null));

        private CommandBase _regenerateApiTokenCommand;
        public CommandBase RegenerateApiTokenCommand => _regenerateApiTokenCommand ?? (_regenerateApiTokenCommand = new DelegateCommand(
            LogManager.GetCurrentClassLogger(),
            _ =>
            {
                if (_messageBox.ConfirmYesNo(GeneralSettingsUI.ApiRegenerateToken_Confirm, GeneralSettingsUI.ApiRegenerateToken_Caption, false))
                {
                    _apiConnection.Regenerate();
                }
            },
            _ => _apiConnection != null));

        private bool _setDpiUnaware;
        public bool SetDpiUnaware
        {
            get => _setDpiUnaware;
            set
            {
                if (_setDpiUnaware != value)
                {
                    _setDpiUnaware = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool SetDpiUnawareEnabled
        {
            get
            {
                var osVersion = _operatingSystem.GetOSVersion();
                return osVersion != null && osVersion >= WindowsVersion.Windows81;
            }
        }

        private bool SynchronizeVbeSettings()
        {
            if (!_messageBox.ConfirmYesNo(GeneralSettingsUI.CompileBeforeParse_WarnCompileOnDemandEnabled,
                GeneralSettingsUI.CompileBeforeParse_WarnCompileOnDemandEnabled_Caption, true))
            {
                return false;
            }

            _vbeSettings.CompileOnDemand = false;
            _vbeSettings.BackGroundCompile = false;
            return true;
        }

        private int _autoSavePeriod;
        public int AutoSavePeriod
        {
            get => _autoSavePeriod;
            set
            {
                if (_autoSavePeriod != value)
                {
                    _autoSavePeriod = value;
                    OnPropertyChanged();
                }
            }
        }

        private DelimiterOptions _delimiter;
        public DelimiterOptions Delimiter
        {
            get => _delimiter;
            set
            {
                if (_delimiter != value)
                {
                    _delimiter = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<MinimumLogLevel> LogLevels { get; set; }
        private MinimumLogLevel _selectedLogLevel;
        private bool _userEditedLogLevel;

        public MinimumLogLevel SelectedLogLevel
        {
            get => _selectedLogLevel;
            set
            {
                if (!Equals(_selectedLogLevel, value))
                {
                    _userEditedLogLevel = true;
                    _selectedLogLevel = value;
                    OnPropertyChanged();
                }
            }
        }

        public CommandBase ShowLogFolderCommand { get; }

        private void ShowLogFolder()
        {
            _operatingSystem.ShowFolder(ApplicationConstants.LOG_FOLDER_PATH);
        }

        public void UpdateConfig(Configuration config)
        {
            config.UserSettings.GeneralSettings = GetCurrentGeneralSettings();
            config.UserSettings.HotkeySettings.Settings = Hotkeys.Select(vm => vm.Unwrap()).ToArray();
        }

        public void SetToDefaults(Configuration config)
        {
            TransferSettingsToView(config.UserSettings.GeneralSettings, config.UserSettings.HotkeySettings);
        }

        private Rubberduck.Settings.GeneralSettings GetCurrentGeneralSettings()
        {
            return new Rubberduck.Settings.GeneralSettings
            {
                Language = SelectedLanguage,
                CanShowSplash = ShowSplashAtStartup,
                CanCheckVersion = CheckVersionAtStartup,
                IncludePreRelease = IncludePreRelease,
                CompileBeforeParse = CompileBeforeParse,
                SetDpiUnaware =  SetDpiUnaware,
                EnableFolderDragAndDrop = EnableFolderDragAndDrop,
                ApiEnabled = ApiEnabled,
                ApiAllowNetwork = ApiAllowNetwork,
                ApiAllowImmediate = ApiAllowImmediate,
                ApiAllowImmediateRead = ApiAllowImmediateRead,
                ApiAllowProcDump = ApiAllowProcDump,
                GitClientEnabled = GitClientEnabled,
                GitPath = GitPath,
                GitAiEnabled = GitAiEnabled,
                GitAiModel = GitAiModel,
                ApiPort = ApiPort,
                IsSmartIndenterPrompted = _indenterPrompted,
                IsAutoSaveEnabled = AutoSaveEnabled,
                AutoSavePeriod = AutoSavePeriod,
                UserEditedLogLevel = _userEditedLogLevel,
                MinimumLogLevel = _selectedLogLevel.Ordinal,
                EnableExperimentalFeatures = ExperimentalFeatures
            };
        }

        protected override void TransferSettingsToView(Rubberduck.Settings.GeneralSettings toLoad)
        {
            TransferSettingsToView(toLoad, null);
        }

        private void TransferSettingsToView(IGeneralSettings general, IHotkeySettings hotkey)
        {
            SelectedLanguage = Languages.FirstOrDefault(culture => culture.Code == general.Language.Code);

            Hotkeys = hotkey == null
                ? new ObservableViewModelCollection<HotkeySettingViewModel>()
                : new ObservableViewModelCollection<HotkeySettingViewModel>(hotkey.Settings.Select(data => new HotkeySettingViewModel(data)));
            ShowSplashAtStartup = general.CanShowSplash;
            CheckVersionAtStartup = general.CanCheckVersion;
            IncludePreRelease = general.IncludePreRelease;
            CompileBeforeParse = general.CompileBeforeParse;
            SetDpiUnaware = general.SetDpiUnaware;
            EnableFolderDragAndDrop = general.EnableFolderDragAndDrop;
            ApiEnabled = general.ApiEnabled;
            ApiAllowNetwork = general.ApiAllowNetwork;
            ApiAllowImmediate = general.ApiAllowImmediate;
            ApiAllowImmediateRead = general.ApiAllowImmediateRead;
            ApiAllowProcDump = general.ApiAllowProcDump;
            GitClientEnabled = general.GitClientEnabled;
            GitPath = general.GitPath;
            GitAiEnabled = general.GitAiEnabled;
            GitAiModel = general.GitAiModel;
            ApiPort = general.ApiPort;
            _indenterPrompted = general.IsSmartIndenterPrompted;
            AutoSaveEnabled = general.IsAutoSaveEnabled;
            AutoSavePeriod = general.AutoSavePeriod;
            _userEditedLogLevel = general.UserEditedLogLevel;
            _selectedLogLevel = LogLevels.First(l => l.Ordinal == general.MinimumLogLevel);

            ExperimentalFeatures = _experimentalFeatureTypes
                .Select(type => {
                    var attribute = (ExperimentalAttribute) type.GetCustomAttributes(typeof(ExperimentalAttribute), false).First();
                    return attribute.Resource;
                })
                .Distinct()
                .Select(resourceKey => new ExperimentalFeature {
                    IsEnabled = general.EnableExperimentalFeatures.SingleOrDefault(d => d.Key == resourceKey)?.IsEnabled ?? false,
                    Key = resourceKey
                })
                .ToList();
        }

        protected override string DialogLoadTitle => SettingsUI.DialogCaption_LoadGeneralSettings;
        protected override string DialogSaveTitle => SettingsUI.DialogCaption_SaveGeneralSettings;

        protected override void ImportSettings()
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = SettingsUI.DialogCaption_LoadGeneralSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;
                var general = Service.Import(dialog.FileName);
                var hotkey = _hotkeyService.Import(dialog.FileName);
                //Always assume Smart Indenter registry import has been prompted if importing.
                general.IsSmartIndenterPrompted = true;
                TransferSettingsToView(general, hotkey);
            }
        }

        protected override void ExportSettings(Rubberduck.Settings.GeneralSettings settings)
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = SettingsUI.DialogMask_XmlFilesOnly,
                Title = SettingsUI.DialogCaption_SaveGeneralSettings
            })
            {
                dialog.ShowDialog();
                if (string.IsNullOrEmpty(dialog.FileName)) return;

                // We call save before export to ensure the UI settings are synced to the service before exporting
                Service.Save(settings);
                _hotkeyService.Save(new HotkeySettings { Settings = Hotkeys.Select(vm => vm.Unwrap()).ToArray() });
                // this assumes Export does not truncate any existing exported settings
                Service.Export(dialog.FileName);
                _hotkeyService.Export(dialog.FileName);
            }
        }
    }
}