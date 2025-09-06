using System;
using System.IO;

using AtualizaAssinaturaOutlook.Modules;
using AtualizaAssinaturaOutlook.Modules.Core;
//using AtualizaAssinaturaOutlook.Modules.Services;



using Microsoft.Extensions.Configuration;

//using static AtualizaAssinaturaOutlook.Modules.ContaPadraoOutlook;
using static AtualizaAssinaturaOutlook.Modules.HtmlSignatureConfigurator;

namespace AtualizaAssinaturaOutlook.Modules
{
	public class SignatureService:ISignatureService
	{
		private readonly IOutlookSignatureProvider _outlookSignatureProvider;
		private readonly IHtmlSignatureConfigurator _htmlSignatureConfigurator;
		private readonly IXmlSignatureConfigurator _xmlSignatureConfigurator;
		private readonly IConfiguration _configuration;

		public SignatureService (
				IOutlookSignatureProvider outlookSignatureProvider,
				IHtmlSignatureConfigurator htmlSignatureConfigurator,
				IXmlSignatureConfigurator xmlSignatureConfigurator,
				IConfiguration configuration )
		{
			_outlookSignatureProvider = outlookSignatureProvider;
			_htmlSignatureConfigurator = htmlSignatureConfigurator;
			_xmlSignatureConfigurator = xmlSignatureConfigurator;
			_configuration = configuration;
		}

		public void RunUpdateProcess ()
		{
			Console.WriteLine ("Iniciando o processo de atualização da assinatura do Outlook...");

			try
			{
				// Obter o nome da assinatura padrão do Outlook
				string defaultSignatureName = _outlookSignatureProvider.GetDefaultSignatureName ();

				if (string.IsNullOrEmpty (defaultSignatureName))
				{
					Console.WriteLine ("Nenhuma assinatura padrão do Outlook encontrada. Processo abortado.");
					return;
				}

				Console.WriteLine ($"Assinatura padrão encontrada: {defaultSignatureName}");

				// Obter o caminho da nova imagem da configuração
				string? newImagePath = _configuration["SignatureSettings:NewImagePath"];
				if (string.IsNullOrEmpty(newImagePath))
				{
					Console.WriteLine("Caminho da nova imagem não configurado em appsettings.json. Usando caminho padrão.");
					newImagePath = "//127.0.0.1/outlook_files/novo_ass_email.jpg"; // Fallback
				}

				// Configurar o arquivo HTML da assinatura
				string htmlFilePath = _htmlSignatureConfigurator.GetSignatureHtmlPath (defaultSignatureName);
				_htmlSignatureConfigurator.ConfigureHtmlImageSource (htmlFilePath, newImagePath!);
				Console.WriteLine ($"Arquivo HTML da assinatura '{defaultSignatureName}.htm' configurado com o novo caminho da imagem.");

				// Configurar o arquivo XML da assinatura
				string xmlOutput = _xmlSignatureConfigurator.ConfigureXmlImageSource (defaultSignatureName, newImagePath!);
				Console.WriteLine ($"Arquivo XML da assinatura '{defaultSignatureName}_arquivos/filelist.xml' configurado com o novo caminho da imagem.");
				// Console.WriteLine("Conteúdo XML atualizado:\n" + xmlOutput);
			}
			catch (FileNotFoundException ex)
			{
				Console.WriteLine ($"Erro: {ex.Message}");
			}
			catch (ArgumentException ex)
			{
				Console.WriteLine ($"Erro de argumento: {ex.Message}");
			}
			catch (Exception ex)
			{
				Console.WriteLine ($"Ocorreu um erro inesperado durante o processo de atualização: {ex.Message}");
				Console.WriteLine (ex.ToString ()); // Log completo da exceção
			}

			Console.WriteLine ("Processo de atualização da assinatura do Outlook concluído.");
		}
	}
}


