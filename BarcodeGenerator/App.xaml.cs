using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace BarcodeGenerator
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        public static bool IsEnglish { get; set; }

        private static string _watchFilePath;

        public static string WatchFilePath
        {
            get { return _watchFilePath; }
            set
            {
                _watchFilePath = value;
                WritePrivateProfileString("Setting", "WatchPath", value, SettingFilePath);
            }
        }

        [DllImport("kernel32.dll")]
        public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder refVal, int size, string filePath);
        [DllImport("kernel32.dll")]
        public static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        public App()
        {
            ErrorHandler.ExceptionHandler.AddHandler(true, false, false);
        }

        private static readonly string SettingFilePath = AppDomain.CurrentDomain.BaseDirectory + @"setting.ini";
        protected override void OnStartup(StartupEventArgs e)
        {
            var parameter = new StringBuilder();
            GetPrivateProfileString("Setting", "Language", "english", parameter, 20, SettingFilePath);
            IsEnglish = parameter.ToString().ToLower() == "english";
            SetLanguage(IsEnglish);
            GetPrivateProfileString("Setting", "WatchPath", "", parameter, 128, SettingFilePath);
            _watchFilePath = parameter.ToString();
            if (string.IsNullOrEmpty(_watchFilePath) || !Directory.Exists(_watchFilePath))
            {
                WatchFilePath = AppDomain.CurrentDomain.BaseDirectory;
            }
            GetPrivateProfileString("Setting", "Flag", "0", parameter, 10, SettingFilePath);
            if (int.TryParse(parameter.ToString(), out var flag) && flag >= 0)
            {
                Flag = flag;
            }
        }

        private static void SetLanguage(bool isEnglish)
        {
            string language = isEnglish ? "English" : "Chinese";
            string requestedCulture = string.Format("Language/{0}.xaml", language);

            try
            {
                ResourceDictionary resourceDictionary = new ResourceDictionary
                {
                    Source = new Uri(requestedCulture, UriKind.Relative)
                };
                int index =
                    Current.Resources.MergedDictionaries.ToList().FindIndex(r => r.Source == resourceDictionary.Source);
                if (index != -1)
                {
                    Current.Resources.MergedDictionaries.RemoveAt(index);
                }
                Current.Resources.MergedDictionaries.Add(resourceDictionary);
            }
            catch (Exception exception)
            {
            }
        }

        public static int Flag { get; set; }
    }
}
