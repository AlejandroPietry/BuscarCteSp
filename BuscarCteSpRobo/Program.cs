using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

var arquivoChaves = @"C:\Users\pietr\Downloads\Chaves Julho 2022.txt";
//var arquivoChaves = @"C:\Users\pietr\Downloads\chaveQueHouveramErros.txt";

var outputDir = @"C:\Users\pietr\OneDrive\Documentos\xmlrobo";

//if (args.Contains("--sample-html", StringComparer.OrdinalIgnoreCase))
//{
//    var sampleHtmlPath = Path.Combine(Directory.GetCurrentDirectory(), "htmldapagina.txt");

//    if (!File.Exists(sampleHtmlPath))
//        throw new FileNotFoundException("Arquivo de HTML de exemplo não encontrado.", sampleHtmlPath);

//    var html = File.ReadAllText(sampleHtmlPath, Encoding.UTF8);
//    var chave = Regex.Match(html, @"DF-e\s*\[(\d{44})\]").Groups[1].Value;
//    var xml = GerarCteProcReconstruido(html, chave);
//    var sampleOutputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample-gerado-refatorado.xml");

//    File.WriteAllText(sampleOutputPath, xml, Encoding.UTF8);
//    Console.WriteLine(sampleOutputPath);
//    return;
//}

if (!File.Exists(arquivoChaves))
    throw new FileNotFoundException("Arquivo de chaves não encontrado.", arquivoChaves);

var chaves = File.ReadLines(arquivoChaves)
    .Select(linha => linha.Trim())
    .Where(linha => !string.IsNullOrWhiteSpace(linha))
    .ToList();

Directory.CreateDirectory(outputDir);

var baseUrl = "https://nfe.fazenda.sp.gov.br/CTeConsulta/Consulta/ExibirConsultaAutenticado";

var options = new ChromeOptions();
options.AddArgument(@"--user-data-dir=C:\SeleniumChromeProfile");
options.AddArgument("--start-maximized");

using var driver = new ChromeDriver(options);
var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

Console.WriteLine("Abra/autentique com certificado se necessário.");
Console.WriteLine("Pressione ENTER para iniciar o processamento...");
Console.ReadLine();

var baixados = 0;
var erros = new List<object>();
var random = new Random();

for (var i = 0; i < chaves.Count; i++)
{
    var chave = chaves[i];
    const string mensagemUfInvalida = "CTE emitida em outra uf";

    try
    {
        Log("======================================");
        Log($"Processando chave {i + 1}/{chaves.Count}: {chave}");

        if (!chave.StartsWith("35"))
        {
            Log($"ERRO na chave {i + 1}/{chaves.Count}: {chave}");
            Log(mensagemUfInvalida);

            erros.Add(new
            {
                Chave = chave,
                Index = i,
                Erro = mensagemUfInvalida,
                DataHora = DateTime.Now,
                Url = ""
            });

            continue;
        }

        var url = $"{baseUrl}?ChaveDFe={chave}&UseRecaptcha=False&TipoConsulta=IMPRESSAO";

        Log("Acessando URL...");
        driver.Navigate().GoToUrl(url);

        Log("Aguardando divImpressao...");
        wait.Until(d => d.FindElement(By.Id("divImpressao")));

        var html = driver.PageSource;

        Log("Gerando cteProc reconstruído...");
        var xml = GerarCteProcReconstruido(html, chave);

        var filePath = Path.Combine(outputDir, $"{chave}.xml");

        Log($"Salvando XML em: {filePath}");
        File.WriteAllText(filePath, xml, Encoding.UTF8);

        baixados++;

        Log("Download/salvamento concluído.");
        Log($"Feito download de {baixados} notas de um total de {chaves.Count}.");

        var delay = 0;
        Log($"Aguardando {delay / 1000.0:N1}s antes da próxima chave...");
        Thread.Sleep(delay);
    }
    catch (Exception ex)
    {
        Log($"ERRO na chave {i + 1}/{chaves.Count}: {chave}");
        Log(ex.Message);

        erros.Add(new
        {
            Chave = chave,
            Index = i,
            Erro = ex.Message,
            DataHora = DateTime.Now,
            Url = driver.Url
        });

        var delay = random.Next(1000, 2000);
        Log($"Erro registrado. Aguardando {delay / 1000.0:N1}s e seguindo...");
        Thread.Sleep(delay);
    }
}

Log("======================================");
Log("PROCESSAMENTO FINALIZADO");
Log($"Total salvo: {baixados}/{chaves.Count}");
Log($"Total de erros: {erros.Count}");

if (erros.Any())
{
    var errosPath = Path.Combine(outputDir, "erros.txt");
    var linhas = erros.Select(e => e.ToString() ?? string.Empty);

    File.WriteAllLines(errosPath, linhas);

    Log($"Arquivo de erros salvo em: {errosPath}");
}

static string GerarCteProcReconstruido(string html, string chaveFallback)
{
    var doc = new HtmlDocument();
    doc.LoadHtml(html);

    var cte = ExtrairCte(doc, chaveFallback);
    return MontarXml(cte);
}

static CTeData ExtrairCte(HtmlDocument doc, string chaveFallback)
{
    var root = doc.GetElementbyId("divImpressao")
        ?? throw new Exception("divImpressao não encontrada.");

    var chave = PegarChave(doc, root, chaveFallback);

    if (string.IsNullOrWhiteSpace(chave) || chave.Length != 44)
        throw new Exception("Chave CT-e não encontrada ou inválida.");

    var principal = PegarSecaoPrincipalPorTitulo(root, "Dados CT-e");
    var emitenteDetalhes = PegarSecaoPrincipalPorTitulo(root, "Dados do Emitente");
    var tomadorDetalhes = PegarSecaoPrincipalPorTitulo(root, "Dados do Tomador (Remetente)", "Dados do Tomador");
    var remetenteDetalhes = PegarSecaoPrincipalPorTitulo(root, "Dados do Remetente");
    var destinatarioDetalhes = PegarSecaoPrincipalPorTitulo(root, "Dados do Destinatario");
    var totais = PegarSecaoPrincipalPorTitulo(root, "Totais");
    var impostos = PegarSecaoPrincipalPorTitulo(root, "Impostos");
    var rodoviario = PegarSecaoPrincipalPorTitulo(root, "Rodoviario");
    var carga = PegarSecaoPorTitulo(root, "Informações da Carga");
    var autorizados = PegarSecaoPorTitulo(root, "Autorizados ao Download");
    var responsavelTecnico = PegarSecaoPorTitulo(root, "Responsável Técnico");
    var suplementares = PegarSecaoPorTitulo(root, "Informações Suplementares");

    var emit = ExtrairParticipanteDetalhado(emitenteDetalhes, "enderEmit", incluiPais: false);
    var toma = ExtrairParticipanteDetalhado(tomadorDetalhes, "enderToma", incluiPais: true);
    var rem = ExtrairParticipanteDetalhado(remetenteDetalhes, "enderReme", incluiPais: true);
    var dest = ExtrairParticipanteDetalhado(destinatarioDetalhes, "enderDest", incluiPais: true);

    var ide = ExtrairIde(principal, chave, emit, rem, dest, toma);
    var vPrest = ExtrairVPrest(totais);
    var imp = ExtrairImpostos(impostos);
    var infCarga = ExtrairCarga(carga);
    emit.Fone = AplicarFallbackTelefoneEmitente(emit.Doc, emit.Fone);
    dest.Fone = AplicarFallbackTelefoneDestinatario(dest.Doc, dest.Fone);

    var prot = ExtrairProtocolo(doc, principal, chave, ide.DhEmi);
    var autXml = ExtrairAutorizados(autorizados);
    var respTec = ExtrairResponsavelTecnico(responsavelTecnico);
    var qrCode = Texto(suplementares, "./table[2]/tbody/tr/td[1]");
    var rntrc = Texto(rodoviario, "./table[2]/tbody/tr[2]/td[1]");

    return new CTeData
    {
        Chave = chave,
        Ide = ide,
        Compl = ExtrairComplementos(ide, imp),
        Emit = emit,
        Rem = rem,
        Dest = dest,
        VPrest = vPrest,
        Imp = imp,
        InfCTeNorm = new InfCteNormData
        {
            InfCarga = infCarga,
            Nfes = ExtrairNfes(carga),
            Rntrc = rntrc
        },
        AutXml = autXml,
        RespTec = respTec,
        QrCode = qrCode,
        Prot = prot
    };
}

static IdeData ExtrairIde(
    HtmlNode? principal,
    string chave,
    ParticipanteData emit,
    ParticipanteData rem,
    ParticipanteData dest,
    ParticipanteData toma)
{
    var tabelaIdentificacao = PegarTabelaPorCabecalhos(principal, "Numero", "Serie", "Data Emissao");
    var tabelaServico = PegarTabelaPorCabecalhos(principal, "Tipo Servico", "Modal", "Finalidade", "Forma");
    var tabelaNatureza = PegarTabelaPorCabecalhos(principal, "Natureza da Prestacao", "CFOP", "Digest Value do CT-e");
    var tabelaTrajeto = PegarTabelaPorCabecalhos(principal, "Inicio da Prestacao", "Fim da Prestacao", "Envio", "Indicador do Tomador");

    var dataEmissao = TextoCelulaTabela(tabelaIdentificacao, 4, 3);
    var envio = SepararUfCidade(TextoCelulaTabela(tabelaTrajeto, 4, 1));
    var tipoServico = SepararCodigoDescricao(TextoCelulaTabela(tabelaServico, 2, 1));
    var modalTexto = TextoCelulaTabela(tabelaServico, 2, 2);
    var modal = SepararCodigoDescricao(modalTexto);
    var finalidade = SepararCodigoDescricao(TextoCelulaTabela(tabelaServico, 2, 3));
    var formaEmissao = SepararCodigoDescricao(TextoCelulaTabela(tabelaServico, 2, 4));
    var indicacaoTomador = SepararCodigoDescricao(TextoCelulaTabela(tabelaTrajeto, 6, 1));
    var tomaCodigo = DeterminarCodigoTomador(toma, emit, rem, dest);
    var municipioEnvioCodigo = ResolverCodigoMunicipio(envio.Uf, envio.Cidade, principal?.OwnerDocument);

    return new IdeData
    {
        CUf = chave[..2],
        CCt = chave.Substring(35, 8),
        Cfop = TextoCelulaTabela(tabelaNatureza, 2, 2),
        NatOp = TextoCelulaTabela(tabelaNatureza, 2, 1),
        Mod = "57",
        Serie = Numeros(TextoCelulaTabela(tabelaIdentificacao, 4, 2)),
        NCt = Numeros(TextoCelulaTabela(tabelaIdentificacao, 4, 1)),
        DhEmi = ConverterDataHoraBrParaIso(dataEmissao),
        TpImp = "1",
        TpEmis = ValorOuPadrao(formaEmissao.Codigo, "1"),
        CDv = chave[^1..],
        TpAmb = "1",
        TpCTe = ValorOuPadrao(finalidade.Codigo, "0"),
        ProcEmi = "0",
        VerProc = "3.00",
        CMunEnv = municipioEnvioCodigo,
        XMunEnv = envio.Cidade,
        UfEnv = envio.Uf,
        Modal = ValorOuPadrao(modal.Codigo, "01"),
        TpServ = ValorOuPadrao(tipoServico.Codigo, "0"),
        CMunIni = rem.Endereco.CMun,
        XMunIni = rem.Endereco.XMun,
        UfIni = rem.Endereco.Uf,
        CMunFim = dest.Endereco.CMun,
        XMunFim = dest.Endereco.XMun,
        UfFim = dest.Endereco.Uf,
        Retira = "0",
        XDetRetira = "NAO SE APLICA",
        IndIeToma = indicacaoTomador.Codigo,
        Toma = tomaCodigo,
        ModalDescricao = modalTexto,
        DigestValue = TextoCelulaTabela(tabelaNatureza, 2, 3)
    };
}

static ParticipanteData ExtrairParticipanteDetalhado(HtmlNode? secao, string enderecoTag, bool incluiPais)
{
    var endereco = ExtrairEndereco(secao, incluiPais);

    if (string.IsNullOrWhiteSpace(endereco.CMun))
        endereco.CMun = ResolverCodigoMunicipio(endereco.Uf, endereco.XMun, secao?.OwnerDocument);

    return new ParticipanteData
    {
        Doc = Numeros(Texto(secao, "./table[2]/tbody/tr[2]/td[1]")),
        XNome = Texto(secao, "./table[2]/tbody/tr[2]/td[2]"),
        Ie = Texto(secao, "./table[2]/tbody/tr[4]/td[1]"),
        XFant = Texto(secao, "./table[2]/tbody/tr[4]/td[2]"),
        Fone = Numeros(Texto(secao, "./table[6]/tbody/tr[2]/td[1]")),
        Email = Texto(secao, "./table[6]/tbody/tr[2]/td[2]"),
        EnderecoTag = enderecoTag,
        Endereco = endereco
    };
}

static EnderecoData ExtrairEndereco(HtmlNode? secao, bool incluiPais)
{
    var enderecoBruto = Texto(secao, "./table[4]/tbody/tr[2]/td[1]");
    var municipioBruto = Texto(secao, "./table[5]/tbody/tr[2]/td[1]");
    var municipio = SepararCodigoDescricao(municipioBruto);
    var enderecoSeparado = SepararEndereco(enderecoBruto);

    return new EnderecoData
    {
        XLgr = enderecoSeparado.Logradouro,
        Nro = enderecoSeparado.Numero,
        XCpl = enderecoSeparado.Complemento,
        XBairro = Texto(secao, "./table[4]/tbody/tr[2]/td[2]"),
        CMun = ValorOuPadrao(
            municipio.Codigo,
            ResolverCodigoMunicipio(Texto(secao, "./table[5]/tbody/tr[2]/td[2]"), municipio.Descricao, secao?.OwnerDocument)),
        XMun = municipio.Descricao,
        Cep = Numeros(Texto(secao, "./table[5]/tbody/tr[2]/td[3]")),
        Uf = Texto(secao, "./table[5]/tbody/tr[2]/td[2]"),
        CPais = incluiPais ? SepararCodigoDescricao(Texto(secao, "./table[5]/tbody/tr[2]/td[4]")).Codigo : "",
        XPais = incluiPais ? SepararCodigoDescricao(Texto(secao, "./table[5]/tbody/tr[2]/td[4]")).Descricao : ""
    };
}

static VPrestData ExtrairVPrest(HtmlNode? secao)
{
        var componentes = secao?.SelectNodes("./table[3]/tbody/tr[position()>1]")
        ?.Select(tr => new PrestacaoComponenteData
        {
            XNome = Limpar(ValorTd(tr, 1)),
            VComp = BrMoneyToXml(ValorTd(tr, 2))
        })
        .Where(comp => !string.IsNullOrWhiteSpace(comp.XNome))
        .ToList() ?? new List<PrestacaoComponenteData>();

    return new VPrestData
    {
        VTPrest = BrMoneyToXml(Texto(secao, "./table[2]/tbody/tr[2]/td[1]")),
        VRec = BrMoneyToXml(Texto(secao, "./table[2]/tbody/tr[2]/td[2]")),
        Componentes = componentes
    };
}

static ImpData ExtrairImpostos(HtmlNode? secao)
{
    return new ImpData
    {
        Cst = Texto(secao, "./table[2]/tbody/tr[2]/td[1]"),
        Vbc = BrMoneyToXml(Texto(secao, "./table[2]/tbody/tr[4]/td[1]")),
        PIcms = PercentualToXml(Texto(secao, "./table[2]/tbody/tr[4]/td[2]")),
        VIcms = BrMoneyToXml(Texto(secao, "./table[2]/tbody/tr[4]/td[3]")),
        VTotTrib = BrMoneyToXml(Texto(secao, "./table[2]/tbody/tr[6]/td[1]"))
    };
}

static InfCargaData ExtrairCarga(HtmlNode? secao)
{
    var quantidades = secao?.SelectNodes("./table[3]/tbody/tr[position()>1]")
        ?.Select(tr =>
        {
            var unidade = SepararCodigoDescricao(ValorTd(tr, 1));

            return new QuantidadeCargaData
            {
                CUnid = unidade.Codigo,
                TpMed = Limpar(ValorTd(tr, 2)),
                QCarga = BrDecimalToXml(ValorTd(tr, 3), 4)
            };
        })
        .Where(item => !string.IsNullOrWhiteSpace(item.CUnid) || !string.IsNullOrWhiteSpace(item.TpMed))
        .ToList() ?? new List<QuantidadeCargaData>();

    return new InfCargaData
    {
        VCarga = BrMoneyToXml(Texto(secao, "./table[2]/tbody/tr[2]/td[1]")),
        ProPred = Texto(secao, "./table[2]/tbody/tr[2]/td[2]"),
        XOutCat = Texto(secao, "./table[2]/tbody/tr[2]/td[3]"),
        Quantidades = quantidades
    };
}

static List<string> ExtrairNfes(HtmlNode? secao)
{
    return secao?.SelectNodes("./table[5]//td")
        ?.Select(td => Numeros(td.InnerText))
        .Where(chave => chave.Length == 44)
        .Distinct()
        .ToList() ?? new List<string>();
}

static ComplementoData ExtrairComplementos(IdeData ide, ImpData imp)
{
    var observacoes = new List<ObservacaoData>();

    if (!string.IsNullOrWhiteSpace(imp.VTotTrib))
    {
        observacoes.Add(new ObservacaoData
        {
            XCampo = "LeidaTransparencia",
            XTexto = $"O valor aproximado de tributos incidentes sobre o preco deste servico e de R$ {XmlMoneyToBr(imp.VTotTrib)}"
        });
    }

    if (!string.IsNullOrWhiteSpace(ide.ModalDescricao))
    {
        var servico = ide.ModalDescricao.StartsWith($"{ide.Modal} - ", StringComparison.OrdinalIgnoreCase)
            ? ide.ModalDescricao
            : !string.IsNullOrWhiteSpace(ide.Modal)
                ? $"{ide.Modal} - {ide.ModalDescricao}"
                : ide.ModalDescricao;

        observacoes.Add(new ObservacaoData
        {
            XCampo = "SERVICO",
            XTexto = RemoverAcentos(servico).ToUpperInvariant()
        });
    }

    return new ComplementoData
    {
        Observacoes = observacoes
    };
}

static ProtocoloData ExtrairProtocolo(HtmlDocument doc, HtmlNode? principal, string chave, string dhEmi)
{
    var modalAutorizacao = doc.GetElementbyId("ev110100");
    var versaoAplicacao = ResolverVerAplicacao(doc, dhEmi);

    return new ProtocoloData
    {
        TpAmb = "1",
        VerAplic = versaoAplicacao,
        ChCTe = chave,
        DhRecbto = ConverterDataHoraBrParaIso(Texto(modalAutorizacao, ".//table[4]/tbody/tr[2]/td[1]")),
        NProt = Numeros(Texto(modalAutorizacao, ".//table[2]/tbody/tr/td[4]")),
        DigVal = Texto(principal, "./table[9]/tbody/tr[2]/td[3]"),
        CStat = "100",
        XMotivo = "Autorizado o uso do CT-e"
    };
}

static List<string> ExtrairAutorizados(HtmlNode? secao)
{
    return secao?.SelectNodes("./table[1]/tr[position()>1]/td[1]")
        ?.Select(td => Numeros(td.InnerText))
        .Where(doc => !string.IsNullOrWhiteSpace(doc))
        .Distinct()
        .ToList() ?? new List<string>();
}

static ResponsavelTecnicoData ExtrairResponsavelTecnico(HtmlNode? secao)
{
    return new ResponsavelTecnicoData
    {
        Cnpj = Numeros(Texto(secao, "./table[1]/tbody/tr[2]/td[1]")),
        XContato = Texto(secao, "./table[1]/tbody/tr[2]/td[2]"),
        Fone = Numeros(Texto(secao, "./table[1]/tbody/tr[4]/td[1]")),
        Email = Texto(secao, "./table[1]/tbody/tr[4]/td[2]")
    };
}

static string MontarXml(CTeData cte)
{
    XNamespace ns = "http://www.portalfiscal.inf.br/cte";

    var cteProc = new XElement(ns + "cteProc",
        new XAttribute("versao", "3.00"),
        new XElement(ns + "CTe",
            new XElement(ns + "infCte",
                new XAttribute("Id", $"CTe{cte.Chave}"),
                new XAttribute("versao", "3.00"),
                MontarIde(ns, cte.Ide),
                MontarCompl(ns, cte.Compl),
                MontarEmit(ns, cte.Emit),
                MontarRem(ns, cte.Rem),
                MontarDest(ns, cte.Dest),
                MontarVPrest(ns, cte.VPrest),
                MontarImp(ns, cte.Imp),
                MontarInfCteNorm(ns, cte.InfCTeNorm),
                cte.AutXml.Select(doc =>
                    new XElement(ns + "autXML",
                        doc.Length == 11
                            ? new XElement(ns + "CPF", doc)
                            : new XElement(ns + "CNPJ", doc)
                    )
                ),
                MontarRespTec(ns, cte.RespTec)
            ),
            new XElement(ns + "infCTeSupl",
                new XElement(ns + "qrCodCTe", cte.QrCode)
            )
        ),
        new XElement(ns + "protCTe",
            new XAttribute("versao", "3.00"),
            new XElement(ns + "infProt",
                new XElement(ns + "tpAmb", cte.Prot.TpAmb),
                ElementoOpcional(ns, "verAplic", cte.Prot.VerAplic),
                new XElement(ns + "chCTe", cte.Prot.ChCTe),
                ElementoOpcional(ns, "dhRecbto", cte.Prot.DhRecbto),
                ElementoOpcional(ns, "nProt", cte.Prot.NProt),
                ElementoOpcional(ns, "digVal", cte.Prot.DigVal),
                new XElement(ns + "cStat", cte.Prot.CStat),
                new XElement(ns + "xMotivo", cte.Prot.XMotivo)
            )
        )
    );

    var xmlDoc = new XDocument(
        new XDeclaration("1.0", "UTF-8", null),
        cteProc
    );

    return xmlDoc.ToString(SaveOptions.DisableFormatting);
}

static XElement MontarIde(XNamespace ns, IdeData ide)
{
    return new XElement(ns + "ide",
        new XElement(ns + "cUF", ide.CUf),
        new XElement(ns + "cCT", ide.CCt),
        ElementoOpcional(ns, "CFOP", ide.Cfop),
        ElementoOpcional(ns, "natOp", ide.NatOp),
        new XElement(ns + "mod", ide.Mod),
        ElementoOpcional(ns, "serie", ide.Serie),
        ElementoOpcional(ns, "nCT", ide.NCt),
        ElementoOpcional(ns, "dhEmi", ide.DhEmi),
        new XElement(ns + "tpImp", ide.TpImp),
        new XElement(ns + "tpEmis", ide.TpEmis),
        new XElement(ns + "cDV", ide.CDv),
        new XElement(ns + "tpAmb", ide.TpAmb),
        new XElement(ns + "tpCTe", ide.TpCTe),
        new XElement(ns + "procEmi", ide.ProcEmi),
        new XElement(ns + "verProc", ide.VerProc),
        ElementoOpcional(ns, "cMunEnv", ide.CMunEnv),
        ElementoOpcional(ns, "xMunEnv", ide.XMunEnv),
        ElementoOpcional(ns, "UFEnv", ide.UfEnv),
        ElementoOpcional(ns, "modal", ide.Modal),
        ElementoOpcional(ns, "tpServ", ide.TpServ),
        ElementoOpcional(ns, "cMunIni", ide.CMunIni),
        ElementoOpcional(ns, "xMunIni", ide.XMunIni),
        ElementoOpcional(ns, "UFIni", ide.UfIni),
        ElementoOpcional(ns, "cMunFim", ide.CMunFim),
        ElementoOpcional(ns, "xMunFim", ide.XMunFim),
        ElementoOpcional(ns, "UFFim", ide.UfFim),
        new XElement(ns + "retira", ide.Retira),
        new XElement(ns + "xDetRetira", ide.XDetRetira),
        ElementoOpcional(ns, "indIEToma", ide.IndIeToma),
        new XElement(ns + "toma3",
            new XElement(ns + "toma", ide.Toma)
        )
    );
}

static object? MontarCompl(XNamespace ns, ComplementoData compl)
{
    if (!compl.Observacoes.Any())
        return null;

    return new XElement(ns + "compl",
        compl.Observacoes.Select(obs =>
            new XElement(ns + "ObsCont",
                new XAttribute("xCampo", obs.XCampo),
                new XElement(ns + "xTexto", obs.XTexto)
            )
        )
    );
}

static XElement MontarEmit(XNamespace ns, ParticipanteData emit)
{
    return new XElement(ns + "emit",
        new XElement(ns + "CNPJ", emit.Doc),
        ElementoOpcional(ns, "IE", emit.Ie),
        ElementoOpcional(ns, "xNome", emit.XNome),
        ElementoOpcional(ns, "xFant", emit.XFant),
        new XElement(ns + "enderEmit",
            ElementoOpcional(ns, "xLgr", emit.Endereco.XLgr),
            ElementoOpcional(ns, "nro", emit.Endereco.Nro),
            ElementoOpcional(ns, "xBairro", emit.Endereco.XBairro),
            ElementoOpcional(ns, "cMun", emit.Endereco.CMun),
            ElementoOpcional(ns, "xMun", emit.Endereco.XMun),
            ElementoOpcional(ns, "CEP", emit.Endereco.Cep),
            ElementoOpcional(ns, "UF", emit.Endereco.Uf),
            ElementoOpcional(ns, "fone", emit.Fone)
        )
    );
}

static XElement MontarRem(XNamespace ns, ParticipanteData rem)
{
    return new XElement(ns + "rem",
        new XElement(ns + "CNPJ", rem.Doc),
        ElementoOpcional(ns, "IE", rem.Ie),
        ElementoOpcional(ns, "xNome", rem.XNome),
        ElementoOpcional(ns, "xFant", rem.XFant),
        ElementoOpcional(ns, "fone", rem.Fone),
        MontarEndereco(ns, rem.EnderecoTag, rem.Endereco, incluirComplemento: false, incluirPais: true),
        ElementoOpcional(ns, "email", rem.Email)
    );
}

static XElement MontarDest(XNamespace ns, ParticipanteData dest)
{
    return new XElement(ns + "dest",
        dest.Doc.Length == 11
            ? new XElement(ns + "CPF", dest.Doc)
            : new XElement(ns + "CNPJ", dest.Doc),
        ElementoOpcional(ns, "xNome", dest.XNome),
        ElementoOpcional(ns, "fone", dest.Fone),
        MontarEndereco(ns, dest.EnderecoTag, dest.Endereco, incluirComplemento: true, incluirPais: true),
        ElementoOpcional(ns, "email", dest.Email)
    );
}

static XElement MontarEndereco(
    XNamespace ns,
    string tagName,
    EnderecoData endereco,
    bool incluirComplemento,
    bool incluirPais)
{
    return new XElement(ns + tagName,
        ElementoOpcional(ns, "xLgr", endereco.XLgr),
        ElementoOpcional(ns, "nro", endereco.Nro),
        incluirComplemento ? ElementoOpcional(ns, "xCpl", endereco.XCpl) : null,
        ElementoOpcional(ns, "xBairro", endereco.XBairro),
        ElementoOpcional(ns, "cMun", endereco.CMun),
        ElementoOpcional(ns, "xMun", endereco.XMun),
        ElementoOpcional(ns, "CEP", endereco.Cep),
        ElementoOpcional(ns, "UF", endereco.Uf),
        incluirPais ? ElementoOpcional(ns, "cPais", endereco.CPais) : null,
        incluirPais ? ElementoOpcional(ns, "xPais", endereco.XPais) : null
    );
}

static XElement MontarVPrest(XNamespace ns, VPrestData dados)
{
    return new XElement(ns + "vPrest",
        ElementoOpcional(ns, "vTPrest", dados.VTPrest),
        ElementoOpcional(ns, "vRec", dados.VRec),
        dados.Componentes.Select(comp =>
            new XElement(ns + "Comp",
                new XElement(ns + "xNome", comp.XNome),
                new XElement(ns + "vComp", comp.VComp)
            )
        )
    );
}

static XElement MontarImp(XNamespace ns, ImpData imp)
{
    return new XElement(ns + "imp",
        new XElement(ns + "ICMS",
            new XElement(ns + "ICMS00",
                ElementoOpcional(ns, "CST", imp.Cst),
                ElementoOpcional(ns, "vBC", imp.Vbc),
                ElementoOpcional(ns, "pICMS", imp.PIcms),
                ElementoOpcional(ns, "vICMS", imp.VIcms)
            )
        ),
        ElementoOpcional(ns, "vTotTrib", imp.VTotTrib)
    );
}

static XElement MontarInfCteNorm(XNamespace ns, InfCteNormData dados)
{
    return new XElement(ns + "infCTeNorm",
        new XElement(ns + "infCarga",
            ElementoOpcional(ns, "vCarga", dados.InfCarga.VCarga),
            ElementoOpcional(ns, "proPred", dados.InfCarga.ProPred),
            ElementoOpcional(ns, "xOutCat", dados.InfCarga.XOutCat),
            dados.InfCarga.Quantidades.Select(q =>
                new XElement(ns + "infQ",
                    ElementoOpcional(ns, "cUnid", q.CUnid),
                    ElementoOpcional(ns, "tpMed", q.TpMed),
                    ElementoOpcional(ns, "qCarga", q.QCarga)
                )
            )
        ),
        new XElement(ns + "infDoc",
            dados.Nfes.Select(nfe =>
                new XElement(ns + "infNFe",
                    new XElement(ns + "chave", nfe)
                )
            )
        ),
        new XElement(ns + "infModal",
            new XAttribute("versaoModal", "3.00"),
            new XElement(ns + "rodo",
                ElementoOpcional(ns, "RNTRC", dados.Rntrc)
            )
        )
    );
}

static object? MontarRespTec(XNamespace ns, ResponsavelTecnicoData resp)
{
    if (string.IsNullOrWhiteSpace(resp.Cnpj))
        return null;

    return new XElement(ns + "infRespTec",
        new XElement(ns + "CNPJ", resp.Cnpj),
        ElementoOpcional(ns, "xContato", resp.XContato),
        ElementoOpcional(ns, "email", resp.Email),
        ElementoOpcional(ns, "fone", resp.Fone)
    );
}

static HtmlNode? PegarSecaoPrincipalPorTitulo(HtmlNode root, params string[] titulosEsperados)
{
    var titulosNormalizados = titulosEsperados
        .Select(NormalizarTituloSecao)
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    return root.SelectNodes("./div")
        ?.FirstOrDefault(secao => titulosNormalizados.Contains(NormalizarTituloSecao(ExtrairTituloSecaoPrincipal(secao))));
}

static HtmlNode? PegarSecaoPorTitulo(HtmlNode root, params string[] titulosEsperados)
{
    var titulosNormalizados = titulosEsperados
        .Select(NormalizarTituloSecao)
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    return root.SelectNodes("./div")
        ?.FirstOrDefault(secao => titulosNormalizados.Contains(NormalizarTituloSecao(ExtrairTituloSecao(secao))));
}

static string ExtrairTituloSecaoPrincipal(HtmlNode? secao)
{
    if (secao is null)
        return "";

    var tituloCabecalho = secao.SelectSingleNode("./table[contains(@class, 'dfe-cabecalho')]//tr/td[2]")
        ?? secao.SelectSingleNode("./table[contains(@class, 'dfe-cabecalho')]//tr/td[normalize-space()]");

    if (tituloCabecalho is not null)
        return Limpar(tituloCabecalho.InnerText);

    return ExtrairTituloSecao(secao);
}

static string ExtrairTituloSecao(HtmlNode? secao)
{
    if (secao is null)
        return "";

    foreach (var child in secao.ChildNodes.Where(node => node.NodeType == HtmlNodeType.Element))
    {
        if (child.Name.Equals("table", StringComparison.OrdinalIgnoreCase))
        {
            var tituloTabela = child.SelectNodes(".//td")
                ?.Select(td => Limpar(td.InnerText))
                .FirstOrDefault(texto => !string.IsNullOrWhiteSpace(texto));

            if (!string.IsNullOrWhiteSpace(tituloTabela))
                return tituloTabela;
        }

        if (child.Name.Equals("div", StringComparison.OrdinalIgnoreCase) &&
            child.GetAttributeValue("class", "").Contains("dfe-titulo", StringComparison.OrdinalIgnoreCase))
        {
            var tituloDiv = Limpar(child.InnerText);

            if (!string.IsNullOrWhiteSpace(tituloDiv))
                return tituloDiv;
        }
    }

    return "";
}

static string NormalizarTituloSecao(string titulo)
{
    return RemoverAcentos(titulo).ToUpperInvariant();
}

static string PegarChave(HtmlDocument doc, HtmlNode div, string fallback)
{
    var h4 = doc.DocumentNode.SelectSingleNode("//h4")?.InnerText;
    var chaveH4 = Numeros(h4);

    if (chaveH4.Length == 44)
        return chaveH4;

    var chaveDiv = Regex.Match(Numeros(div.InnerText), @"\d{44}").Value;

    if (chaveDiv.Length == 44)
        return chaveDiv;

    return fallback;
}

static string Texto(HtmlNode? contexto, string xpath)
{
    return Limpar(contexto?.SelectSingleNode(xpath)?.InnerText);
}

static HtmlNode? PegarTabelaPorCabecalhos(HtmlNode? secao, params string[] cabecalhosEsperados)
{
    if (secao is null)
        return null;

    var cabecalhosNormalizados = cabecalhosEsperados
        .Select(NormalizarTituloSecao)
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    return secao.SelectNodes("./table")
        ?.FirstOrDefault(tabela =>
        {
            var cabecalhosTabela = tabela
                .SelectNodes("./thead/tr/th | ./tbody/tr/th | ./tr/th")
                ?.Select(th => NormalizarTituloSecao(th.InnerText))
                .Where(texto => !string.IsNullOrWhiteSpace(texto))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return cabecalhosTabela is not null && cabecalhosNormalizados.All(cabecalhosTabela.Contains);
        });
}

static string TextoCelulaTabela(HtmlNode? tabela, int linha, int coluna)
{
    var row = tabela?.SelectSingleNode($"./tbody/tr[{linha}] | ./tr[{linha}]");
    return ValorTd(row, coluna);
}

static string ValorTd(HtmlNode? row, int indice)
{
    return Limpar(row?.SelectSingleNode($"./td[{indice}]")?.InnerText);
}

static (string Codigo, string Descricao) SepararCodigoDescricao(string texto)
{
    var limpo = Limpar(texto);
    var match = Regex.Match(limpo, @"^(?<codigo>\d+)\s*-\s*(?<descricao>.+)$");

    if (match.Success)
        return (match.Groups["codigo"].Value, match.Groups["descricao"].Value);

    return ("", limpo);
}

static (string Uf, string Cidade) SepararUfCidade(string texto)
{
    var limpo = Limpar(texto);
    var match = Regex.Match(limpo, @"^(?<uf>[A-Z]{2})\s*-\s*(?<cidade>.+)$");

    if (match.Success)
        return (match.Groups["uf"].Value, match.Groups["cidade"].Value);

    return ("", limpo);
}

static (string Logradouro, string Numero, string Complemento) SepararEndereco(string texto)
{
    var partes = Limpar(texto)
        .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

    if (partes.Length == 0)
        return ("", "", "");

    if (partes.Length == 1)
        return (partes[0], "", "");

    if (partes.Length == 2)
        return (partes[0], partes[1], "");

    return (partes[0], partes[1], string.Join(", ", partes.Skip(2)));
}

static string ConverterDataHoraBrParaIso(string texto)
{
    var limpo = Limpar(texto).Replace("às", "-");
    var match = Regex.Match(limpo, @"(?<data>\d{2}/\d{2}/\d{4}).*?(?<hora>\d{2}:\d{2}:\d{2}).*?(?<offset>[+-]\d{2}:\d{2})");

    if (!match.Success)
        return limpo;

    var dataHoraBr = $"{match.Groups["data"].Value} {match.Groups["hora"].Value}";

    if (!DateTime.TryParseExact(
        dataHoraBr,
        "dd/MM/yyyy HH:mm:ss",
        CultureInfo.InvariantCulture,
        DateTimeStyles.None,
        out var parsed))
    {
        return limpo;
    }

    return $"{parsed:yyyy-MM-ddTHH:mm:ss}{match.Groups["offset"].Value}";
}

static string BrMoneyToXml(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
        return "";

    var limpo = Limpar(value)
        .Replace("%", "")
        .Replace(".", "")
        .Replace(",", ".");

    return decimal.TryParse(limpo, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
        ? parsed.ToString("0.00", CultureInfo.InvariantCulture)
        : limpo;
}

static string PercentualToXml(string? value)
{
    return BrMoneyToXml(value);
}

static string BrDecimalToXml(string? value, int casas)
{
    if (string.IsNullOrWhiteSpace(value))
        return "";

    var limpo = Limpar(value)
        .Replace(".", "")
        .Replace(",", ".");

    return decimal.TryParse(limpo, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
        ? parsed.ToString($"0.{new string('0', casas)}", CultureInfo.InvariantCulture)
        : limpo;
}

static string XmlMoneyToBr(string? value)
{
    if (string.IsNullOrWhiteSpace(value))
        return "";

    return decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
        ? parsed.ToString("N2", new CultureInfo("pt-BR"))
        : value;
}

static string Numeros(string? txt)
{
    return Regex.Replace(Limpar(txt), @"\D", "");
}

static string Limpar(string? txt)
{
    return Regex.Replace(
        HtmlEntity.DeEntitize(txt ?? "")
        .Replace('\u00A0', ' ')
        .Replace('Â', ' ')
        .Replace('Ã', ' ')
        .Replace(" ", " "),
        @"\s+",
        " ")
        .Trim();
}

static string RemoverAcentos(string texto)
{
    var normalized = Limpar(texto).Normalize(NormalizationForm.FormD);
    var builder = new StringBuilder();

    foreach (var ch in normalized)
    {
        if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
            builder.Append(ch);
    }

    return builder.ToString().Normalize(NormalizationForm.FormC);
}

static string ValorOuPadrao(string valor, string padrao)
{
    return string.IsNullOrWhiteSpace(valor) ? padrao : valor;
}

static string ExtrairVersaoAplicacao(HtmlDocument doc)
{
    var texto = Limpar(doc.DocumentNode.InnerText);
    var match = Regex.Match(texto, @"Vers[aã]o\s+(SP-CTe-[0-9\-]+)");

    return match.Success ? match.Groups[1].Value : "";
}

static string ResolverVerAplicacao(HtmlDocument doc, string dhEmi)
{
    var versaoExtraida = ExtrairVersaoAplicacao(doc);

    if (string.IsNullOrWhiteSpace(versaoExtraida))
        return FallbackDados.VerAplicPadraoHistorico;

    if (!TryParseVersaoAplicacaoDate(versaoExtraida, out var dataVersao))
        return versaoExtraida;

    if (!TryParseIsoDate(dhEmi, out var dataEmissao))
        return versaoExtraida;

    return dataVersao.Date > dataEmissao.Date
        ? FallbackDados.VerAplicPadraoHistorico
        : versaoExtraida;
}

static bool TryParseVersaoAplicacaoDate(string verAplic, out DateTime data)
{
    var match = Regex.Match(verAplic, @"SP-CTe-(\d{4})-(\d{2})-(\d{2})-\d+");

    if (!match.Success)
    {
        data = default;
        return false;
    }

    return DateTime.TryParseExact(
        $"{match.Groups[1].Value}-{match.Groups[2].Value}-{match.Groups[3].Value}",
        "yyyy-MM-dd",
        CultureInfo.InvariantCulture,
        DateTimeStyles.None,
        out data);
}

static bool TryParseIsoDate(string isoDate, out DateTime data)
{
    if (DateTimeOffset.TryParse(isoDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
    {
        data = parsed.DateTime;
        return true;
    }

    data = default;
    return false;
}

static string AplicarFallbackTelefoneEmitente(string cnpj, string telefone)
{
    if (!string.IsNullOrWhiteSpace(telefone))
        return telefone;

    return FallbackDados.TelefonesEmitentePorCnpj.TryGetValue(cnpj, out var fallback)
        ? fallback
        : telefone;
}

static string AplicarFallbackTelefoneDestinatario(string documento, string telefone)
{
    if (!string.IsNullOrWhiteSpace(telefone) && telefone.Length >= 12)
        return telefone;

    return FallbackDados.TelefonesDestinatarioPorDocumento.TryGetValue(documento, out var fallback)
        ? fallback
        : telefone;
}

static string ResolverCodigoMunicipio(string uf, string cidade, HtmlDocument? doc)
{
    if (string.IsNullOrWhiteSpace(uf) || string.IsNullOrWhiteSpace(cidade))
        return "";

    var cidadeNormalizada = RemoverAcentos(cidade).ToUpperInvariant();
    var htmlTexto = doc?.DocumentNode.InnerText ?? "";
    var regex = new Regex($@"(?<codigo>\d{{7}})\s*-\s*{Regex.Escape(cidadeNormalizada)}", RegexOptions.IgnoreCase);
    var match = regex.Match(RemoverAcentos(htmlTexto).ToUpperInvariant());

    if (match.Success)
        return match.Groups["codigo"].Value;

    return DadosFixos.MunicipiosConhecidos.TryGetValue($"{uf}|{cidadeNormalizada}", out var codigo)
        ? codigo
        : "";
}

static string DeterminarCodigoTomador(
    ParticipanteData toma,
    ParticipanteData emit,
    ParticipanteData rem,
    ParticipanteData dest)
{
    if (!string.IsNullOrWhiteSpace(toma.Doc))
    {
        if (toma.Doc == rem.Doc)
            return "0";

        if (toma.Doc == dest.Doc)
            return "3";

        if (toma.Doc == emit.Doc)
            return "4";
    }

    return "0";
}

static XElement? ElementoOpcional(XNamespace ns, string nome, string? valor)
{
    return string.IsNullOrWhiteSpace(valor) ? null : new XElement(ns + nome, valor);
}

static void Log(string msg)
{
    Console.WriteLine($"[CTE PROC BOT] {DateTime.Now:HH:mm:ss} - {msg}");
}

static class DadosFixos
{
    public static readonly Dictionary<string, string> MunicipiosConhecidos = new(StringComparer.OrdinalIgnoreCase)
    {
        ["SP|SAO PAULO"] = "3550308"
    };
}

static class FallbackDados
{
    public const string VerAplicPadraoHistorico = "SP-CTe-2021-08-19-1";

    public static readonly Dictionary<string, string> TelefonesEmitentePorCnpj = new(StringComparer.OrdinalIgnoreCase)
    {
        ["20147617002276"] = "551121216161"
    };

    public static readonly Dictionary<string, string> TelefonesDestinatarioPorDocumento = new(StringComparer.OrdinalIgnoreCase)
    {
        ["07116763562"] = "071992954393"
    };
}

sealed class CTeData
{
    public string Chave { get; set; } = "";
    public IdeData Ide { get; set; } = new();
    public ComplementoData Compl { get; set; } = new();
    public ParticipanteData Emit { get; set; } = new();
    public ParticipanteData Rem { get; set; } = new();
    public ParticipanteData Dest { get; set; } = new();
    public VPrestData VPrest { get; set; } = new();
    public ImpData Imp { get; set; } = new();
    public InfCteNormData InfCTeNorm { get; set; } = new();
    public List<string> AutXml { get; set; } = new();
    public ResponsavelTecnicoData RespTec { get; set; } = new();
    public string QrCode { get; set; } = "";
    public ProtocoloData Prot { get; set; } = new();
}

sealed class IdeData
{
    public string CUf { get; set; } = "";
    public string CCt { get; set; } = "";
    public string Cfop { get; set; } = "";
    public string NatOp { get; set; } = "";
    public string Mod { get; set; } = "";
    public string Serie { get; set; } = "";
    public string NCt { get; set; } = "";
    public string DhEmi { get; set; } = "";
    public string TpImp { get; set; } = "";
    public string TpEmis { get; set; } = "";
    public string CDv { get; set; } = "";
    public string TpAmb { get; set; } = "";
    public string TpCTe { get; set; } = "";
    public string ProcEmi { get; set; } = "";
    public string VerProc { get; set; } = "";
    public string CMunEnv { get; set; } = "";
    public string XMunEnv { get; set; } = "";
    public string UfEnv { get; set; } = "";
    public string Modal { get; set; } = "";
    public string TpServ { get; set; } = "";
    public string CMunIni { get; set; } = "";
    public string XMunIni { get; set; } = "";
    public string UfIni { get; set; } = "";
    public string CMunFim { get; set; } = "";
    public string XMunFim { get; set; } = "";
    public string UfFim { get; set; } = "";
    public string Retira { get; set; } = "";
    public string XDetRetira { get; set; } = "";
    public string IndIeToma { get; set; } = "";
    public string Toma { get; set; } = "";
    public string ModalDescricao { get; set; } = "";
    public string DigestValue { get; set; } = "";
}

sealed class ComplementoData
{
    public List<ObservacaoData> Observacoes { get; set; } = new();
}

sealed class ObservacaoData
{
    public string XCampo { get; set; } = "";
    public string XTexto { get; set; } = "";
}

sealed class ParticipanteData
{
    public string Doc { get; set; } = "";
    public string XNome { get; set; } = "";
    public string Ie { get; set; } = "";
    public string XFant { get; set; } = "";
    public string Fone { get; set; } = "";
    public string Email { get; set; } = "";
    public string EnderecoTag { get; set; } = "";
    public EnderecoData Endereco { get; set; } = new();
}

sealed class EnderecoData
{
    public string XLgr { get; set; } = "";
    public string Nro { get; set; } = "";
    public string XCpl { get; set; } = "";
    public string XBairro { get; set; } = "";
    public string CMun { get; set; } = "";
    public string XMun { get; set; } = "";
    public string Cep { get; set; } = "";
    public string Uf { get; set; } = "";
    public string CPais { get; set; } = "";
    public string XPais { get; set; } = "";
}

sealed class VPrestData
{
    public string VTPrest { get; set; } = "";
    public string VRec { get; set; } = "";
    public List<PrestacaoComponenteData> Componentes { get; set; } = new();
}

sealed class PrestacaoComponenteData
{
    public string XNome { get; set; } = "";
    public string VComp { get; set; } = "";
}

sealed class ImpData
{
    public string Cst { get; set; } = "";
    public string Vbc { get; set; } = "";
    public string PIcms { get; set; } = "";
    public string VIcms { get; set; } = "";
    public string VTotTrib { get; set; } = "";
}

sealed class InfCteNormData
{
    public InfCargaData InfCarga { get; set; } = new();
    public List<string> Nfes { get; set; } = new();
    public string Rntrc { get; set; } = "";
}

sealed class InfCargaData
{
    public string VCarga { get; set; } = "";
    public string ProPred { get; set; } = "";
    public string XOutCat { get; set; } = "";
    public List<QuantidadeCargaData> Quantidades { get; set; } = new();
}

sealed class QuantidadeCargaData
{
    public string CUnid { get; set; } = "";
    public string TpMed { get; set; } = "";
    public string QCarga { get; set; } = "";
}

sealed class ResponsavelTecnicoData
{
    public string Cnpj { get; set; } = "";
    public string XContato { get; set; } = "";
    public string Email { get; set; } = "";
    public string Fone { get; set; } = "";
}

sealed class ProtocoloData
{
    public string TpAmb { get; set; } = "";
    public string VerAplic { get; set; } = "";
    public string ChCTe { get; set; } = "";
    public string DhRecbto { get; set; } = "";
    public string NProt { get; set; } = "";
    public string DigVal { get; set; } = "";
    public string CStat { get; set; } = "";
    public string XMotivo { get; set; } = "";
}
