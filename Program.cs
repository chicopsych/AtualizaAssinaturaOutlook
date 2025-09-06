using System;

using AtualizaAssinaturaOutlook.Modules;
using AtualizaAssinaturaOutlook.Modules.Core;
//using AtualizaAssinaturaOutlook.Modules.Services;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.CommandLine;
using Microsoft.Extensions.DependencyInjection;

using static AtualizaAssinaturaOutlook.Modules.HtmlSignatureConfigurator;

namespace AtualizaAssinaturaOutlook.Modules
{
	internal class Program
	{
		static void Main ( string[] args )
		{
			// 1. Configuração
			var configuration = new ConfigurationBuilder ()
					.AddJsonFile ("appsettings.json", optional: true, reloadOnChange: true)
					.AddCommandLine (args)
					.Build ();

			// 2. Configuração de Serviços (Injeção de Dependência)
			var serviceProvider = new ServiceCollection ()
					.AddSingleton<IConfiguration> (configuration)
					.AddTransient<ISignatureService, SignatureService> ()
					.AddTransient<IOutlookSignatureProvider, OutlookSignatureRegistryProvider> ()
					.AddTransient<IHtmlSignatureConfigurator, HtmlSignatureConfigurator> ()
					.AddTransient<IXmlSignatureConfigurator, XmlSignatureConfigurator> ()
					.BuildServiceProvider ();

			// 3. Execução da Aplicação
			try
			{
				var signatureService = serviceProvider.GetService<ISignatureService> ();
				if (signatureService != null)
				{
					signatureService.RunUpdateProcess ();
				}
				else
				{
					Console.WriteLine ("Serviço de assinatura não foi encontrado.");
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine ($"Ocorreu um erro inesperado: {ex.Message}");
				// Logar o erro completo (ex.ToString())
			}

			Console.WriteLine ("Pressione qualquer tecla para sair...");
			Console.ReadKey ();
		}
	}
}
