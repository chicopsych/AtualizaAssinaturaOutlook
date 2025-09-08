using System;
using System.Runtime.Versioning;

using Microsoft.Win32;
using System.IO;
using AtualizaAssinaturaOutlook.Modules.Core;

namespace AtualizaAssinaturaOutlook.Modules
{
	public interface IOutlookSignatureProvider
	{
		string GetDefaultSignatureName ();
	}

	public class OutlookSignatureRegistryProvider:IOutlookSignatureProvider
	{
		private readonly string[] _officeVersions = { "16.0", "15.0", "14.0", "13.0", "12.0" };

		public string GetDefaultSignatureName ()
		{
			// Verifica se está rodando no Windows antes de chamar métodos restritos
			if (!OperatingSystem.IsWindows())
			{
				throw new PlatformNotSupportedException("Este método só é suportado no Windows.");
			}

			foreach (string version in _officeVersions)
			{
				// 1. Tenta o caminho direto (Common\\MailSettings)
				string? signature = TryGetSignatureFromMailSettings (version);
				if (!string.IsNullOrEmpty (signature))
				{
					return signature;
				}

				// 2. Fallback: perfis do Outlook
				signature = TryGetSignatureFromProfiles (version);
				if (!string.IsNullOrEmpty (signature))
				{
					return signature;
				}
			}

			return string.Empty; // Nenhuma assinatura padrão encontrada

		}

		[SupportedOSPlatform("windows")]
		private string? TryGetSignatureFromMailSettings ( string version )
		{
			try
			{
				using (var mailSettings = Registry.CurrentUser.OpenSubKey ($@"Software\Microsoft\Office\{version}\Common\MailSettings"))
				{
					return mailSettings?.GetValue ("NewSignature") as string ?? string.Empty;
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine ($"Erro ao tentar ler MailSettings para a versão{version}: {ex.Message}");
				return string.Empty;
			}
		}

		[SupportedOSPlatform("windows")]
		private string TryGetSignatureFromProfiles ( string version )
		{
			string profilesPath = $@"Software\Microsoft\Office\{version}\Outlook\Profiles";
			try
			{
				using (RegistryKey? profilesKey = Registry.CurrentUser.OpenSubKey(profilesPath))
				{
					if (profilesKey == null) return string.Empty; // Corrigido para retornar string.Empty

					foreach (string profileName in profilesKey.GetSubKeyNames())
					{
						using (RegistryKey? profileKey = profilesKey.OpenSubKey(Path.Combine(profilesPath, profileName)))
						{
							if (profileKey != null)
							{
								string signature = FindSignatureInProfile(profileKey);

								if (!string.IsNullOrEmpty(signature))
								{
									return signature;
								}
							}
						}
					}
				}

			}
			catch (Exception ex)
			{
				Console.WriteLine ($"Erro ao tentar ler perfis para a versão {version}: {ex.Message}");
			}
			return string.Empty; // Corrigido para retornar string.Empty
		}

		private string FindSignatureInProfile ( RegistryKey profileKey )
		{
			// Garante que nunca retorna nulo, apenas string.Empty se não houver valor
			// Adiciona atributo para garantir que só será chamado em Windows
			[System.Runtime.Versioning.SupportedOSPlatform("windows")]
			string GetSignature(RegistryKey key)
			{
				return key.GetValue("New Signature") as string ?? string.Empty;
			}

			// Adiciona verificação de plataforma antes de chamar método restrito
			if (!OperatingSystem.IsWindows())
			{
				throw new PlatformNotSupportedException("Este método só é suportado no Windows.");
			}
			return GetSignature(profileKey);
		}
	}
}
