using DSharpPlus;
using DSharpPlus.CommandsNext;
using DSharpPlus.SlashCommands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yakumo.Commands.Slash;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using IWshRuntimeLibrary;
using System.Reflection;

namespace Yakumo
{
    internal class Program
    {
        private static DiscordClient Client { get; set; }

        static void CreateShortcut()
        {
            string exePath = Assembly.GetExecutingAssembly().Location;
            string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupFolderPath, "Yakumo.lnk");

            if (System.IO.File.Exists(shortcutPath)) return;

            WshShell wshShell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)wshShell.CreateShortcut(shortcutPath);
            shortcut.Description = "wanna make a vow of mastery?";
            shortcut.TargetPath = exePath;
            shortcut.WorkingDirectory = Path.GetDirectoryName(exePath);
            shortcut.Save();
        }

        static async Task Main(string[] args)
        {
            CreateShortcut();

            var discordConfig = new DiscordConfiguration()
            {
                Intents = DiscordIntents.All,
                Token = "Your discord bot token here",
                TokenType = TokenType.Bot,
                AutoReconnect = true
            };

            Client = new DiscordClient(discordConfig);

            Client.Ready += Client_Ready;

            var slashCommandsConfig = Client.UseSlashCommands();
            slashCommandsConfig.RegisterCommands<SlashCommands>();

            await Client.ConnectAsync();
            await Task.Delay(-1);
        }

        private static Task Client_Ready(DiscordClient sender, DSharpPlus.EventArgs.ReadyEventArgs args)
        {
            return Task.CompletedTask;
        }
    }
}