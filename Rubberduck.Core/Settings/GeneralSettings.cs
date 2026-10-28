using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Xml.Serialization;
using Rubberduck.Common;
using Rubberduck.Resources;

namespace Rubberduck.Settings
{
    public interface IGeneralSettings 
    {
        DisplayLanguageSetting Language { get; set; }
        bool CanShowSplash { get; set; }
        bool CanCheckVersion { get; set; }
        string ApiBaseUrl { get; set; }
        bool IncludePreRelease { get; set; }
        bool CompileBeforeParse { get; set; }
        bool IsSmartIndenterPrompted { get; set; }
        bool IsAutoSaveEnabled { get; set; }
        int AutoSavePeriod { get; set; }
        bool UserEditedLogLevel { get; set; }
        int MinimumLogLevel { get; set; }
        bool SetDpiUnaware { get; set; }
        bool EnableFolderDragAndDrop { get; set; }
        List<ExperimentalFeature> EnableExperimentalFeatures { get; set; }
        bool ApiEnabled { get; set; }
        bool ApiAllowNetwork { get; set; }
        bool ApiAllowImmediate { get; set; }
        bool ApiAllowImmediateRead { get; set; }
        int ApiPort { get; set; }
    }

    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    [XmlType(AnonymousType = true)]
    public class GeneralSettings : IGeneralSettings, IEquatable<GeneralSettings>
    {
        private DisplayLanguageSetting _language = new DisplayLanguageSetting(Locales.DefaultCulture.Name);

        public DisplayLanguageSetting Language
        {
            get => _language;
            set
            {
                if (Locales.AvailableCultures.Exists(culture => culture.Name.Equals(value.Code, StringComparison.OrdinalIgnoreCase)))
                {
                    _language = value;
                }
            }
        }

        public bool CanShowSplash { get; set; }
        public bool CanCheckVersion { get; set; }
        public string ApiBaseUrl { get; set; }
        public bool IncludePreRelease { get; set; }
        public bool CompileBeforeParse { get; set; }
        public bool IsSmartIndenterPrompted { get; set; }
        public bool IsAutoSaveEnabled { get; set; }
        public int AutoSavePeriod { get; set; }

        public bool UserEditedLogLevel { get; set; }

        private int _logLevel;
        public int MinimumLogLevel
        {
            get => _logLevel;
            set
            {
                if (value < LogLevelHelper.MinLogLevel())
                {
                    _logLevel = LogLevelHelper.MinLogLevel();
                }
                else if (value > LogLevelHelper.MaxLogLevel())
                {
                    _logLevel = LogLevelHelper.MaxLogLevel();
                }
                else
                {
                    _logLevel = value;
                }               
            }
        }

        public bool SetDpiUnaware { get; set; }
        public bool EnableFolderDragAndDrop { get; set; }

        public List<ExperimentalFeature> EnableExperimentalFeatures { get; set; } = new List<ExperimentalFeature>();

        /// <summary>Lets AI agents read and edit modules over HTTP. Off until the user switches it on.</summary>
        public bool ApiEnabled { get; set; }

        /// <summary>Answer on the machine's network addresses, not only on loopback, for agents that run in a sandbox or a VM.</summary>
        public bool ApiAllowNetwork { get; set; }

        /// <summary>Lets agents read the Immediate window and enter commands in it, which runs them. Off until the user switches it on.</summary>
        public bool ApiAllowImmediate { get; set; }

        /// <summary>Lets agents read what the Immediate window shows (<c>Debug.Print</c> output) without being able to run anything. Implied by <see cref="ApiAllowImmediate"/>. Off until the user switches it on.</summary>
        public bool ApiAllowImmediateRead { get; set; }

        public int ApiPort { get; set; }

        public GeneralSettings()
        {
            //Enforce non-default default value for members
            //In other words, if we want a bool to default to
            //true, it must be set here for correct behavior
            CompileBeforeParse = true;
            ApiPort = 7437;
        }

        public bool Equals(GeneralSettings other)
        {
            return other != null &&
                   Language.Equals(other.Language) &&
                   CanShowSplash == other.CanShowSplash &&
                   CanCheckVersion == other.CanCheckVersion &&
                   IncludePreRelease == other.IncludePreRelease &&
                   CompileBeforeParse == other.CompileBeforeParse &&
                   IsSmartIndenterPrompted == other.IsSmartIndenterPrompted &&
                   IsAutoSaveEnabled == other.IsAutoSaveEnabled &&
                   AutoSavePeriod == other.AutoSavePeriod &&
                   UserEditedLogLevel == other.UserEditedLogLevel &&
                   MinimumLogLevel == other.MinimumLogLevel &&                   
                   EnableExperimentalFeatures.Count == other.EnableExperimentalFeatures.Count &&
                   EnableExperimentalFeatures.All(other.EnableExperimentalFeatures.Contains) &&
                   SetDpiUnaware == other.SetDpiUnaware &&
                   EnableFolderDragAndDrop == other.EnableFolderDragAndDrop &&
                   ApiEnabled == other.ApiEnabled &&
                   ApiAllowNetwork == other.ApiAllowNetwork &&
                   ApiAllowImmediate == other.ApiAllowImmediate &&
                   ApiAllowImmediateRead == other.ApiAllowImmediateRead &&
                   ApiPort == other.ApiPort &&
                   ApiBaseUrl == other.ApiBaseUrl;
        }
    }
}