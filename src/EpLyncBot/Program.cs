using System;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using NLog;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types;
using System.Runtime.InteropServices;
using System.Threading;

namespace EpLyncBot
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;


        static ITelegramBotClient _bot;
        static LyncConversationHandler _lyncConversationHandler;

        static void Main(string[] args)
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            Settings.Build();

            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var exitEvent = new System.Threading.ManualResetEvent(false);
            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                exitEvent.Set();
            };
            var correlationId = Guid.NewGuid();
            var log = LogManager.GetCurrentClassLogger();
            log.Info("Starting...");
            try
            {
                _bot = new TelegramBotClient(Settings.App.Token);

                _bot.OnMessage += _onMessage;

                _lyncConversationHandler = new LyncConversationHandler(_bot);

                _bot.StartReceiving();

                log.Info("Running");
                exitEvent.WaitOne();

                log.Info("Stopping...");
                _bot.StopReceiving();
            }
            catch (Exception ex)
            {
                log.Fatal(ex);
            }

            log.Info("Stopped");
        }

        static void _onMessage(object sender, MessageEventArgs messageEventArgs)
        {
            var log = LogManager.GetCurrentClassLogger();

            try
            {
                Message message = messageEventArgs.Message;

                if (message.ReplyToMessage != null)
                {
                    var text = message.ReplyToMessage.Text;
                    var nlIndex = text.LastIndexOf('\n');
                    var replyTo = text.Substring(nlIndex + 1);

                    _lyncConversationHandler.SendMessage(replyTo, message.Text);
                }

            }
            catch (Exception ex)
            {
                log.Error(ex);
                _bot.SendTextMessageAsync(new ChatId(Settings.App.ChatId), ex.Message);
            }
        }
    }
}
