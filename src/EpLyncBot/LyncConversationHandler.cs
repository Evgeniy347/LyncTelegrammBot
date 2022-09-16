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


            string[] targetContactUris = { "sip:emarkin@wss-consulting.ru" };
            LyncClient client = LyncClient.GetClient();
            Conversation conv = client.ConversationManager.AddConversation();

            foreach (string target in targetContactUris)
            {
                conv.AddParticipant(client.ContactManager.GetContactByUri(target));
            }
            InstantMessageModality m = conv.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality;
            m.BeginSendMessage("Test Message", null, null);


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
            foreach (Conversation conversation in _lync.ConversationManager.Conversations)
                _onConversationAdded(conversation);

            IsInit = true;
        }

        void _onConversationAdded(object sender, ConversationManagerEventArgs ea)
        {
            _onConversationAdded(ea.Conversation);
        }

        void _onConversationAdded(Conversation conversation)
        {
            System.Console.WriteLine("conversation added");

            _conversations.Add(Guid.NewGuid().ToString(), conversation);
            conversation.ParticipantAdded += _onParticipantAdded;
            conversation.ParticipantRemoved += _onParticipantRemoved;

            foreach (Participant participant in conversation.Participants)
                _onParticipantAdded(participant);
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
            _onParticipantAdded(ea.Participant); 
        }

        void _onParticipantAdded(Participant participant)
        {
            System.Console.WriteLine("part added");

            // if (ea.Participant.IsSelf) return;
            EventHandler<MessageSentEventArgs> onMessageReceived = (s, e) => _onMessageReceived(s, e, participant);

            (participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived -= onMessageReceived;
            (participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived += onMessageReceived;
        }

        void _onParticipantRemoved(object sender, ParticipantCollectionChangedEventArgs ea)
        {
            System.Console.WriteLine("part removed");

            (ea.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality).InstantMessageReceived -=
                (s, e) => _onMessageReceived(s, e, ea.Participant);
        }

        async void _onMessageReceived(object sender, MessageSentEventArgs e, Participant p)
        {
            try
            {
                System.Console.WriteLine("message received");

                if (p.IsSelf && _isSessionLock) return;

                var conversation = (sender as InstantMessageModality).Conversation;

                if (!_conversations.ContainsValue(conversation))
                {
                    _log.Warn("Conversation not found");
                    return;
                }

                var id = _conversations.First(x => x.Value == conversation).Key;

                var message = new StringBuilder();
                message.Append($"*{p.Contact.GetContactInformation(ContactInformationType.DisplayName)}*\n");
                message.Append(_escapeMarkdown(e.Text));
                if (p.IsSelf) message.Append("\n");
                message.Append(id);

                await _bot.SendTextMessageAsync(new ChatId(Settings.App.ChatId), message.ToString(), ParseMode.Markdown, disableNotification: p.IsSelf);
                _log.Info($"sent: {message.ToString()}");
            }
            catch (Exception ex)
            {
                _log.Error(ex);
            }
        }

        string _escapeMarkdown(string str)
        {
            if (string.IsNullOrEmpty(str)) return null;

            string[] markdownChars = { "_", "*" };

            foreach (var c in markdownChars)
            {
                for (var i = str.IndexOf(c, 0); i != -1; i = str.IndexOf(c, i + 2))
                {
                    str = str.Insert(i, "\\");
                }
            }

            return str;
        }
    }
}