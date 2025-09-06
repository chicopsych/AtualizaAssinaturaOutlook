# 📨 Atualização Automática da Imagem de Assinatura do Outlook

Mantém a imagem (logo / banner) da assinatura de e‑mail do Outlook de todos os usuários sempre atualizada apenas trocando um único arquivo em um compartilhamento de rede.

---

## ✨ Visão Geral

Este utilitário em .NET automatiza a atualização da imagem usada na assinatura padrão do Outlook (cliente desktop). Ele:

1. Detecta a assinatura padrão registrada no Windows (varre múltiplas versões do Office).  
2. Localiza o arquivo HTML da assinatura em `%AppData%/Microsoft/Signatures`.  
3. Substitui o atributo `src` do primeiro `<img>` pelo caminho UNC configurado.  
4. Ajusta o `filelist.xml` (quando existe) para garantir que referências à imagem (`image001.jpg`) apontem para o novo recurso.  
5. Permite override de configuração via linha de comando (útil em scripts corporativos / GPO).  

Ideal para cenários onde a equipe de marketing troca periodicamente um banner de campanha sem precisar distribuir manualmente novas assinaturas para todos os colaboradores.

---

## 🧩 Arquitetura & Componentes

| Componente | Responsabilidade |
|------------|------------------|
| `Program` | Monta DI (injeção de dependências) e dispara o processo. |
| `ISignatureService` / `SignatureService` | Orquestra o fluxo completo de atualização. |
| `IOutlookSignatureProvider` / `OutlookSignatureRegistryProvider` | Lê do Registro do Windows qual é a assinatura padrão (procura em várias versões do Office). |
| `IHtmlSignatureConfigurator` / `HtmlSignatureConfigurator` | Carrega e modifica o arquivo `.htm` da assinatura (troca o `src` da imagem). |
| `IXmlSignatureConfigurator` / `XmlSignatureConfigurator` | Atualiza o `filelist.xml` dentro da pasta `*_arquivos` quando presente. |
| `appsettings.json` | Armazena a configuração `SignatureSettings:NewImagePath`. |
| Linha de comando | Pode sobrescrever qualquer chave de configuração (ex.: `--SignatureSettings:NewImagePath=...`). |

Fluxo simplificado:

```text
[Registry] -> Nome da assinatura -> [%AppData%/Microsoft/Signatures]
 |--> Atualiza assinatura.htm (img src)
 |--> Atualiza *_arquivos/filelist.xml (HRef)
```

---

## ✅ Requisitos

- Windows (necessário para acesso ao Registro e estrutura de assinaturas do Outlook).  
- Outlook instalado e pelo menos uma assinatura criada.  
- .NET SDK 9.0 ou superior (para compilar/rodar).  
- Permissão de leitura no Registro do usuário (`HKCU`).  
- Permissão de escrita em `%AppData%/Microsoft/Signatures`.  
- Caminho UNC acessível para a imagem (ex.: `//servidor/marketing/assinatura/banner.jpg`).  

---

## ⚙️ Configuração

Arquivo `appsettings.json` padrão:

```json
{
 "SignatureSettings": {
  "NewImagePath": "//127.0.0.1/outlook_files/novo_ass_email.jpg"
 }
}
```

Sobrescrever via linha de comando (exemplo):

```pwsh
dotnet run -- --SignatureSettings:NewImagePath=//fileserver/public/assinaturas/campanha_natal.jpg
```

Também é possível empacotar o binário e executar com um atalho de logon / GPO apontando o argumento desejado.

---

## ▶️ Execução Rápida

Clonar e executar:

```pwsh
git clone <url-do-repositorio>
cd AtualizaAssinaturaOutlook
dotnet build
dotnet run
```

Publicar (para distribuir um executável self-contained — opcional):

```pwsh
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained true
```

O executável ficará em: `bin/Release/net9.0/win-x64/publish/`.

---

## 🧪 Compatibilidade

| Componente | Suporte | Observação |
|------------|---------|------------|
| Outlook 2016/365 (16.0) | ✅ | Alvo principal/testado |
| Outlook 2013 (15.0) | ⚠️ | Pode funcionar; não testado recentemente |
| Outlook 2010 ou anterior | ❌ | Fora de escopo |
| Windows (x64) | ✅ | Requer acesso ao Registro HKCU |
| Linux / macOS | ❌ | Depende de Outlook Windows e Registro |

---

## 🔍 Detalhes do Processo

1. Varredura de versões do Office: `16.0`, `15.0`, `14.0`, `13.0`, `12.0`.  
2. Tenta chave `HKCU\Software\Microsoft\Office\<ver>\Common\MailSettings` (`NewSignature`).  
3. Fallback: percorre perfis em `...\Outlook\Profiles`.  
4. Localiza `AssinaturaNome.htm`.  
5. Atualiza `<img src="..."></img>`.  
6. Ajusta `filelist.xml` substituindo referências (quando existir).  
7. Registra mensagens no console.  

---

## ⚠️ Limitações Conhecidas

- Necessita que uma assinatura já exista (não cria estrutura do zero).  
- Só altera o primeiro elemento `<img>` encontrado (simplificação deliberada).  
- Não replica a imagem localmente; depende do caminho UNC estar acessível quando o Outlook renderiza.  
- Não personaliza dados dinâmicos (nome/cargo) – foco apenas na imagem/bandeira.  
- `filelist.xml` fora do padrão (namespaces diferentes) pode exigir ajustes futuros.  

---

## � Considerações de Rede / Acesso

- Garanta que o caminho UNC use permissões somente leitura para usuários finais.  
- Se houver autenticação por credenciais, monte previamente a sessão (script de logon).  
- Para resiliência, pode-se futuramente copiar a imagem para `%AppData%` e referenciar localmente (não implementado).  

---

## �🛡️ Tratamento de Erros (Exemplos)

| Situação | Ação |
|----------|------|
| Assinatura não encontrada | Processo é abortado com aviso. |
| Arquivo `.htm` ausente | Lança `FileNotFoundException`. |
| `filelist.xml` inexistente | Ignora silenciosamente a etapa XML. |
| Caminho da imagem vazio | Lança `ArgumentException`. |
| Ambiente não Windows | Lança `PlatformNotSupportedException`. |

---

## 🧪 Ideias de Validação Interna

- Verificar se a imagem aparece em novo e-mail no Outlook.  
- Abrir o `.htm` atualizado em um navegador para inspecionar o `<img src>` final.  
- Confirmar que o caminho UNC é alcançável pelo usuário (testar via Explorer).  

---

## 🛠️ Automação Corporativa (Sugestão)

- Publicar binário em pasta compartilhada.  
- Criar Script PowerShell de logon que execute o utilitário com o parâmetro atualizado.  
- Marketing altera somente o arquivo de imagem no compartilhamento.  
- Usuários recebem a nova imagem no próximo logon / próxima execução programada (Task Scheduler).  

Exemplo de tarefa agendada (PowerShell linha de comando simplificada):

```pwsh
Start-Process "C:\CorpTools\AtualizaAssinaturaOutlook.exe" --SignatureSettings:NewImagePath=//intranet/branding/banner_atual.jpg -WindowStyle Hidden
```

---

## ❓ Por que não um Add-in do Outlook?

Uma abordagem via add-in (VSTO ou Web Add-in) adicionaria complexidade de deployment, atualização e possíveis prompts de segurança. Este utilitário atua só no sistema de arquivos e Registro de usuário, reduzindo atrito operacional e evitando dependência da API de composição do Outlook.

---

## 🧩 Solução de Problemas

| Problema | Possível Causa | Solução |
|----------|----------------|---------|
| Assinatura não muda | Nome incorreto detectado ou cache do Outlook | Fechar e reabrir Outlook; validar nome em `%AppData%/Microsoft/Signatures`. |
| Imagem quebrada | Caminho UNC inacessível | Testar acesso via Explorer; checar permissões de rede. |
| Não encontra `<img>` | Template da assinatura sem imagem inicial | Inserir manualmente uma imagem inicial e rodar novamente. |
| XML não altera | Estrutura fora do padrão | Verificar conteúdo de `filelist.xml` e namespaces; adaptar código se necessário. |
| Erro de plataforma | Rodando fora do Windows | Executar somente em máquinas Windows com Outlook. |

---

## 🗺️ Roadmap (Ideias Futuras)

- [ ] Suporte a múltiplas imagens / variantes por departamento.  
- [ ] Template base de assinatura com placeholders (nome / cargo).  
- [ ] Logs estruturados (Serilog) + nível configurável.  
- [ ] Testes unitários (mocks de Registry e FileSystem).  
- [ ] Validação de reachability da imagem antes de aplicar.  
- [ ] Modo "dry-run" (mostra o que seria alterado).  
- [ ] Geração opcional de assinatura alternativa para respostas.  

---

## 🤝 Contribuindo

1. Faça um fork.  
2. Crie uma branch: `feat/minha-ideia`.  
3. Commit: `git commit -m "feat: adiciona X"`.  
4. Abra um Pull Request descrevendo a motivação.  

Sugestões de melhorias são bem-vindas.

---

## 📄 Licença

Licenciado sob a [MIT License](./LICENSE.md).

---

## ⚠️ Aviso

Este projeto modifica arquivos de assinatura do usuário. Recomenda-se testar em ambiente de homologação antes de adoção ampla.

---

## 📬 Contato

Abra uma issue com dúvidas ou ideias.

---

Se desejar, posso gerar também uma versão em inglês ou adicionar badges de CI. Basta pedir.
