using Microsoft.Lync.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Timers;

namespace LyncAimpUpdater
{
    class Program
    {
        private const string FileName = "CurrentTrackInfo.txt";
        private const string AimpProcessName = "AIMP3";
        private const string Prefix = "Listening to: ";

        private static bool _running;
        private static string _lastStatus;
        private static Timer _timer;

        static void Main()
        {
            _running = true;

            _timer = new Timer(5000);
            _timer.Elapsed += TimerOnElapsed;
            _timer.Start();

            var watcher = new FileSystemWatcher
            {
                Path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), 
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = FileName,
                EnableRaisingEvents = true
            };
            watcher.Changed += WatcherOnChanged;

            while (_running)
            {
                watcher.WaitForChanged(WatcherChangeTypes.Changed);
            }
        }

        private static void WatcherOnChanged(object sender, FileSystemEventArgs fileSystemEventArgs)
        {
            var tempStatus = _lastStatus;
            try
            {
                var client = LyncClient.GetClient();
                if (client == null || client.State != ClientState.SignedIn)
                    return;

                var text = File.ReadAllText(fileSystemEventArgs.FullPath);
                if (String.IsNullOrWhiteSpace(text))
                    return;

                text = Prefix + text.Trim();

                if (_lastStatus == text)
                    return;

                _lastStatus = text;

                client.Self.BeginPublishContactInformation(new[]
                {
                    new KeyValuePair<PublishableContactInformationType, object>(PublishableContactInformationType.PersonalNote, text)
                }, x =>
                {
                    if (x.IsCompleted)
                    {
                        client.Self.EndPublishContactInformation(x);
                    }
                }, DateTime.Now.Ticks);
            }
            catch
            {
                _lastStatus = tempStatus;
            }
        }

        private static void TimerOnElapsed(object sender, ElapsedEventArgs elapsedEventArgs)
        {
            if (_lastStatus == string.Empty)
                return;

            if (IsAimpRunning())
                return;

            var tempStatus = _lastStatus;
            try
            {
                var client = LyncClient.GetClient();
                if (client == null || client.State != ClientState.SignedIn)
                    return;

                _lastStatus = string.Empty;

                client.Self.BeginPublishContactInformation(new[]
                    {
                        new KeyValuePair<PublishableContactInformationType, object>(PublishableContactInformationType.PersonalNote, string.Empty)
                    }, x =>
                    {
                        if (x.IsCompleted)
                        {
                            client.Self.EndPublishContactInformation(x);
                        }
                    }, DateTime.Now.Ticks);
            }
            catch
            {
                _lastStatus = tempStatus;
            }
        }

        private static bool IsAimpRunning()
        {
            return Process.GetProcesses().Any(x => x.ProcessName == AimpProcessName);
        }
    }
}