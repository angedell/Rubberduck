using System;
using NLog;
using Rubberduck.Common;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;

namespace Rubberduck.Api
{
    /// <summary>
    /// <see cref="ILogControl"/> on Rubberduck's settings and NLog. The level is saved through the settings service, which tells the rest of the add-in,
    /// and applied directly as well, so the answer does not depend on who is listening.
    /// </summary>
    public class RubberduckLogControl : ILogControl
    {
        public const string LogFileName = "RubberduckLog.txt";

        private readonly IConfigurationService<Configuration> _settings;

        public RubberduckLogControl(IConfigurationService<Configuration> settings)
        {
            _settings = settings;
        }

        public string ConfiguredLevel => LogLevel.FromOrdinal(_settings.Read().UserSettings.GeneralSettings.MinimumLogLevel).Name;

        //Off is not a threshold NLog holds on to by itself: logging can be disabled outright while the threshold still reads as before.
        public string EffectiveLevel => LogManager.IsLoggingEnabled() ? LogManager.GlobalThreshold.Name : LogLevel.Off.Name;

        public string LogFile => System.IO.Path.Combine(ApplicationConstants.LOG_FOLDER_PATH, LogFileName);

        public void SetLevel(string level)
        {
            var wanted = LogLevel.FromString(level);
            var configuration = _settings.Read();

            //The user's own choice: otherwise a Trace level is taken for the installer's default and reset to Off when Rubberduck shuts down.
            configuration.UserSettings.GeneralSettings.MinimumLogLevel = wanted.Ordinal;
            configuration.UserSettings.GeneralSettings.UserEditedLogLevel = true;
            _settings.Save(configuration);

            LogLevelHelper.SetMinimumLogLevel(wanted);
        }
    }
}
