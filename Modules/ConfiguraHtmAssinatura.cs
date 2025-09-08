using System;
using System.IO;
using System.Xml;

using HtmlAgilityPack;


namespace AtualizaAssinaturaOutlook.Modules
{
	public interface IHtmlSignatureConfigurator
	{
		string GetSignatureHtmlPath ( string signatureName );
		void ConfigureHtmlImageSource ( string htmlFilePath, string newImagePath );
	}

	public class HtmlSignatureConfigurator:IHtmlSignatureConfigurator
	{
		public string GetSignatureHtmlPath ( string signatureName )
		{
			if (string.IsNullOrWhiteSpace (signatureName))
			{
				throw new ArgumentException ("Nome da assinatura não pode ser nulo ou vazio.", nameof (signatureName));
			}

			string appDataPath = Environment.GetFolderPath (Environment.SpecialFolder.ApplicationData);
			string signatureDir = Path.Combine (appDataPath, "Microsoft", "Signatures");
			string signaturePathFile = Path.Combine (signatureDir, $"{signatureName}.htm");

			if (!File.Exists (signaturePathFile))
			{
				throw new FileNotFoundException ($"Arquivo HTM da assinatura não encontrado: {signaturePathFile}");
			}
			return signaturePathFile;
		}

		public void ConfigureHtmlImageSource ( string htmlFilePath, string newImagePath )
		{
			if (string.IsNullOrWhiteSpace (htmlFilePath))
			{
				throw new ArgumentException ("Caminho do arquivo HTML não pode ser nulo ou vazio.", nameof (htmlFilePath));
			}
			if (!File.Exists (htmlFilePath))
			{
				throw new FileNotFoundException ($"Arquivo HTML não encontrado: {htmlFilePath}");
			}
			if (string.IsNullOrWhiteSpace (newImagePath))
			{
				throw new ArgumentException ("Novo caminho da imagem não pode ser nulo ou vazio.", nameof (newImagePath));
			}

			var docHtm = new HtmlDocument ();
			docHtm.Load (htmlFilePath);

			var imgNode = docHtm.DocumentNode.SelectSingleNode ("//img");
			if (imgNode != null)
			{
				imgNode.SetAttributeValue ("src", newImagePath);
				docHtm.Save (htmlFilePath);
			}
			else
			{
				Console.WriteLine ($"Nenhum nó <img> encontrado no arquivo HTML: {htmlFilePath}");

			}

		}

		public interface IXmlSignatureConfigurator
		{
			string ConfigureXmlImageSource ( string baseName, string newImagePath );

		}

		public class XmlSignatureConfigurator:IXmlSignatureConfigurator
		{
			public string ConfigureXmlImageSource ( string baseName, string newImagePath )
			{
				if (string.IsNullOrWhiteSpace (baseName))
				{
					throw new ArgumentException ("O nome base da assinatura não pode ser nulo ou vazio.", nameof (baseName));
				}
				if (string.IsNullOrWhiteSpace (newImagePath))
				{
					throw new ArgumentException ("O novo caminho da imagem não pode ser nulo ou vazio.", nameof (newImagePath));
				}

				string dataPath = Environment.GetFolderPath (Environment.SpecialFolder.ApplicationData);
				string signatureFolder = Path.Combine (dataPath, "Microsoft", "Signatures", $"{baseName}_arquivos");
				string xmlPath = Path.Combine (signatureFolder, "filelist.xml");

				if (!File.Exists (xmlPath))
				{
					throw new FileNotFoundException ($"Arquivo XML não encontrado: {xmlPath}");
				}

				var doc = new XmlDocument ();
				doc.Load (xmlPath);

				var nsmgr = new XmlNamespaceManager (doc.NameTable);
				nsmgr.AddNamespace ("o", "urn:schemas-microsoft-com:office:office");

				XmlNodeList? fileNodes = doc.SelectNodes ("//o:File", nsmgr);
				if (fileNodes == null || fileNodes.Count == 0)
				{
					// Fallback para quando o XML não usa o namespace declarado
					fileNodes = doc.GetElementsByTagName ("File");
					if (fileNodes == null || fileNodes.Count == 0)
					{
						fileNodes = doc.GetElementsByTagName ("o:File"); // Tenta novamente com o prefixo
					}

				}
				bool changed = false;
				foreach (XmlNode fileNode in fileNodes)
				{
					var hrefAttr = fileNode.Attributes?["HRef"];
					if (hrefAttr != null && !string.IsNullOrEmpty (hrefAttr.Value) && hrefAttr.Value.IndexOf ("image001.jpg", StringComparison.OrdinalIgnoreCase) >= 0)
					{

						hrefAttr.Value = newImagePath;
						changed = true;

					}

				}
				if (changed)
				{
					doc.Save (xmlPath);
				}
				return doc.OuterXml;

			}

		}
	}
}
