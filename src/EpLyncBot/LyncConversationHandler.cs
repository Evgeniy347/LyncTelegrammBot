using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Win32;
using NLog;
using Telegram.Bot;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;

namespace EpLyncBot
{
    public class LyncConversationHandler
    {
        readonly ITelegramBotClient _bot;
        readonly Logger _log = LogManager.GetCurrentClassLogger();
        Dictionary<string, Conversation> _conversations;
        bool _isSessionLock = false;
        LyncClient _lync;

        public bool IsInit = false;

        public LyncConversationHandler(ITelegramBotClient bot)
        {
            _bot = bot;
            _initClient();

            Microsoft.Win32.SystemEvents.SessionSwitch += (s, e) =>
            {
                if (e.Reason == SessionSwitchReason.SessionLock)
                {
                    _isSessionLock = true;
                }
                else if (e.Reason == SessionSwitchReason.SessionUnlock)
                {
                    _isSessionLock = false;
                }
            };
        }

        public void SendMessage(string conversationId, string message)
        {
            if (!_conversations.ContainsKey(conversationId)) throw new KeyNotFoundException("Conversation not found");

            var modality = _conversations[conversationId].Modalities[ModalityTypes.InstantMessage] as InstantMessageModality;
            modality.BeginSendMessage(message, null, null);
        }

        void _initClient()
        {
            IsInit = false;

            while (_lync == null)
            {
                try
                {
                    _lync = LyncClient.GetClient();

                    if (_lync.State != ClientState.SignedIn)
                    {
                        _lync = null;
                        Thread.Sleep(1000);
                        _log.Debug("client is not signed in");
                    }
                }
                catch (ClientNotFoundException)
                {
                    Thread.Sleep(1000);
                    _log.Debug("client not found");
                }
            }

            _lync.StateChanged += (s, e) =>
            {
                System.Console.WriteLine(e.NewState);
                if (e.NewState != ClientState.SignedIn)
                {
                    _lync = null;
                    _initClient();
                }
            };

            _conversations = new Dictionary<string, Conversation>();

            _lync.ConversationManager.ConversationAdded += _onConversationAdded;
            _lync.ConversationManager.ConversationRemoved += _onConversationRemoved;

            IsInit = true;
        }

        void _onConversationAdded(object sender, ConversationManagerEventArgs ea)
        {
            System.Console.WriteLine("conversation added");

            _conversations.Add(Guid.NewGuid().ToString(), ea.Conversation);
            ea.Conversation.ParticipantAdded += _onParticipantAdded;
            ea.Conversation.ParticipantRemoved += _onParticipantRemoved;
        }

        void _onConversationRemoved(object sender, ConversationManagerEventArgs e)
        {
            System.Console.Write("conv removed " + _conversations.Count);

            e.Conversation.ParticipantAdded -= _onParticipantAdded;
            e.Conversation.ParticipantRemoved -= _onParticipantRemoved;

            if (_conversations.ContainsValue(e.Conversation))
            {
                _conversations.Remove(_conversations.First(x => x.Value == e.Conversation).Key);
            }
            System.Console.WriteLine(" " + _conversations.Count);
        }

        void _onParticipantAdded(object sender, ParticipantCollectionChangedEventArgs ea)
        {
            System.Console.WriteLine("part added");

            if (ea.Participant.IsSelf) return;

            (ea.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived +=
                (s, e) => _onMessageReceived(s, e, ea.Participant);
        }

        void _onParticipantRemoved(object sender, ParticipantCollectionChangedEventArgs ea)
        {
            System.Console.WriteLine("part removed");

            (ea.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived -=
                (s, e) => _onMessageReceived(s, e, ea.Participant);
        }

        async void _onMessageReceived(object sender, MessageSentEventArgs e, Participant p)
        {
            System.Console.WriteLine("message received");

            var conversation = (sender as InstantMessageModality).Conversation;

            if (!_conversations.ContainsValue(conversation))
            {
                _log.Warn("Conversation not found");
                return;
            }

            var id = _conversations.First(x => x.Value == conversation).Key;

            var message = new StringBuilder();
            message.Append($"*{p.Contact.GetContactInformation(ContactInformationType.DisplayName)}*\n");
            message.Append(e.Text);
            message.Append(id);

            await _bot.SendTextMessageAsync(new ChatId(Settings.App.ChatId), message.ToString(), ParseMode.Markdown, disableNotification: !_isSessionLock);
        }
    }
}