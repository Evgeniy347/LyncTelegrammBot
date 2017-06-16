using Microsoft.Extensions.Configuration;
using NLog.Config;
using NLog;

namespace EpLyncBot
{
	public static class Settings
	{
		public static void Build()
		{
			var builder = new ConfigurationBuilder()
				.AddJsonFile("settings.json", false, true)
				.Build();

			builder.GetSection("App").Bind(App = new AppSettings());

			LogManager.Configuration = new XmlLoggingConfiguration("nlog.config");
		}

		public static AppSettings App { get; private set; }

		public class AppSettings
		{
            public string Token { get; set; }
            public long ChatId { get; set; }
		}
	}
}
