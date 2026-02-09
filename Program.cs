using System.Net;
using System.Xml.Linq;
using ClosedXML.Excel;
using System.Text.Json;
using System.Security.Cryptography.X509Certificates;

class Estado
{
    public int codigo_uf { get; set; }
    public string? uf { get; set; }
}
class Municipio
{
    public int codigo_ibge { get; set; }
    public string? nome { get; set; }
    public int codigo_uf { get; set; }
}
class NotaFiscal
{
    public long NumeroNota { get; set; }
    public DateTime DataEmissao { get; set; }
    public string? Nome { get; set; }
    public string? Valor { get; set; }
    public string? Aliquota { get; set; }
    public string? EstMun { get; set; }
    public string? ItemServ{ get; set; }
    public string? ChaveAcesso { get; set;}
    public bool ValidTomador { get; set;} = true;
    public bool Cancelada {get; set;} = false;
}
class Program
{
    static List<Municipio> municipios = new List<Municipio>();
    static async Task Main(string[] args)
    {
        try
        {
            GatherMunicipio();
            string path = CreateDirectory();    

            Console.WriteLine("Selecione um tipo de nota:\nP - Serviços Prestados.  T - Serviços Tomados");
            string typeInput;
            while(true)
            {    
                typeInput = Console.ReadLine() ?? "";
                if(!typeInput.Equals("P", StringComparison.OrdinalIgnoreCase) && !typeInput.Equals("T", StringComparison.OrdinalIgnoreCase)) Console.WriteLine("Tipo inexistente.");
                else break;
            }
            string[] xmlFiles = Directory.GetFiles(path, "*.xml");
            var processedXML = await ProcessXML(xmlFiles, typeInput);
            string sheetPath = GenerateSheet(processedXML.OrderBy(x => x.DataEmissao.Date).ThenBy(x => x.NumeroNota).ToList());

            Console.WriteLine($"Planilha salva como {sheetPath}");
        }
        catch(Exception ex)
        {
            Console.WriteLine($"Ocorreu um erro: {ex.Message} em {ex.StackTrace}");
            Console.ReadKey();
        }
        
    }
    static string CreateDirectory()
    {
        const string folderName = "ArquivosXML";

        if(!Directory.Exists(folderName))
            Directory.CreateDirectory(folderName);
        return folderName;
    }
    static async Task<List<NotaFiscal>> ProcessXML(string[] xmls, string typeInput)
    {
        if(xmls.Length == 0) throw new Exception("Nenhum arquivo na pasta. Tente novamete...");

        //Begin connection with API 
        X509Certificate2 cert = X509CertificateLoader.LoadPkcs12FromFile(
            @"F:\G2Ka\CERTIFICADOS A1\ASSOCIACAO GESTAO VEICULAR UNIVERSO_14777297000100.pfx",
            "123456"
        );
        var handler = new HttpClientHandler();
        handler.ClientCertificates.Add(cert);
        using HttpClient client = new HttpClient(handler);
        
        List<NotaFiscal> nfs = new List<NotaFiscal>();
        foreach(var xml in xmls)
        {
            XDocument doc = XDocument.Load(xml);
            XNamespace nf = "http://www.sped.fazenda.gov.br/nfse";

            var nfseInfo = doc.Descendants(nf + "infNFSe").FirstOrDefault();
            var emitInfo = doc.Descendants(nf + "emit").FirstOrDefault();
            var tomaInfo = doc.Descendants(nf + "toma").FirstOrDefault();
            var valInfo = doc.Descendants(nf + "vServPrest").FirstOrDefault();
            var valTribInfo = doc.Descendants(nf + "valores").FirstOrDefault();
            var tribInfo = doc.Descendants(nf + "tribMun").FirstOrDefault();
            var canc = doc.Descendants(nf + "NfseCancelamento").FirstOrDefault();
            var serv = doc.Descendants(nf + "cServ").FirstOrDefault();
            
            long numNota = long.Parse(nfseInfo?.Element(nf + "nNFSe")?.Value.ToString() ?? "0");
            DateTime dataEmi = DateTime.Parse(nfseInfo?.Element(nf + "dhProc")?.Value ?? "");
            string nome;
            string estMun = "";
            if(typeInput.Equals("P", StringComparison.OrdinalIgnoreCase))
            {
                nome = tomaInfo?.Element(nf + "xNome")?.Value ?? "";
                string munXML = tomaInfo?.Descendants(nf + "endNac")?.FirstOrDefault()?.Element(nf + "cMun")?.Value ?? "";
                estMun = municipios.First(x => x.codigo_ibge.ToString() == munXML).nome ?? throw new Exception("Município Inválido");
            }
            else
            {
                nome = emitInfo?.Element(nf + "xNome")?.Value ?? "";
                string munXML = emitInfo?.Descendants(nf + "enderNac")?.FirstOrDefault()?.Element(nf + "cMun")?.Value ?? "";
                estMun = municipios.First(x => x.codigo_ibge.ToString() == munXML).nome ?? throw new Exception("Município Inválido");
            }
            string valor = valInfo?.Element(nf + "vServ")?.Value?.Replace(".", ",") ?? "";

            bool issRetido = tribInfo?.Element(nf + "tpRetISSQN")?.Value == "2";
            string? aliquota = tribInfo?.Element(nf + "pAliq")?.Value?.Replace(".", ",");
            string? aliquotaAplic = valTribInfo?.Element(nf + "pAliqAplic")?.Value?.Replace(".", ",");
            string aliquotaIss;
            if(typeInput.Equals("T",StringComparison.OrdinalIgnoreCase))
            {
                if(!string.IsNullOrEmpty(aliquota))
                    aliquotaIss = issRetido ? aliquota : "0,00";
                else
                    aliquotaIss = issRetido ? aliquotaAplic ?? "0" : "0,00";    
            }
            else
            {
                aliquotaIss = !string.IsNullOrEmpty(aliquotaAplic) ? aliquotaAplic ?? "0" : aliquota ?? "0";
            }
            
            string codServ = serv?.Element(nf + "cTribNac")?.Value?.Substring(0, 4).Insert(2, ".") ?? throw new Exception("Código de serviço inválido.");
            
            bool sitCanc = false;
            string chave = nfseInfo?.Attribute("Id")?.Value.Substring(3) ?? throw new Exception("Chave de acesso inválida.");
            if(canc == null)
                sitCanc = await VerifyStatus(chave, client);
            else
                sitCanc = true;

            bool tomador = tomaInfo != null;
            
            NotaFiscal notaObj = new NotaFiscal
            {
                NumeroNota = numNota,
                DataEmissao = dataEmi,
                Nome = nome,
                Valor = valor,
                Aliquota = aliquotaIss,
                ItemServ = codServ,
                EstMun = estMun,
                ChaveAcesso = xml,
                Cancelada = sitCanc,
                ValidTomador = tomador
            };
            nfs.Add(notaObj);
        }
        return nfs;
    }

    static string GenerateSheet(List<NotaFiscal> notas)
    {
        using(var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Relatório de Notas Processadas");

            ws.Row(1).Cell("A").Value = "Cancelada";
            ws.Row(1).Cell("A").Style.Fill.BackgroundColor = XLColor.CoralRed;
            ws.Row(1).Cell("B").Value = "Sem Tomador";
            ws.Row(1).Cell("B").Style.Fill.BackgroundColor = XLColor.Yellow;
            ws.Row(2).Style.Fill.BackgroundColor = XLColor.LightGray;
            ws.Row(1).Style.Font.Bold = true;
            ws.Row(2).Style.Font.Bold = true;
            ws.Row(2).Cell("A").Value = "Sequencia";
            ws.Row(2).Cell("B").Value = "Data de Emissão";
            ws.Row(2).Cell("C").Value = "Número nota";
            ws.Row(2).Cell("D").Value = "Nome";
            ws.Row(2).Cell("E").Value = "Local";
            ws.Row(2).Cell("F").Value = "Cod. Serv";
            ws.Row(2).Cell("G").Value = "Valor Bruto";
            ws.Row(2).Cell("H").Value = "Alíquota";
            ws.Row(2).Cell("I").Value = "Valor ISSQN";
            ws.Row(2).Cell("J").Value = "Chave Acesso";

            int x = 3;
            for(int i = 0; i < notas.Count(); i++)
            {
                if(notas[i].Cancelada) ws.Row(x).Style.Fill.BackgroundColor = XLColor.CoralRed;
                else if(!notas[i].ValidTomador) ws.Row(x).Style.Fill.BackgroundColor = XLColor.Yellow;
                ws.Row(x).Cell("A").Value =  x - 2;
                ws.Row(x).Cell("B").Value = notas[i].DataEmissao.ToString("dd/MM/yyyy");
                ws.Row(x).Cell("C").Value = notas[i].NumeroNota.ToString();
                ws.Row(x).Cell("D").Value = notas[i].Nome;
                ws.Row(x).Cell("E").Value = notas[i].EstMun;
                ws.Row(x).Cell("F").Value = notas[i].ItemServ;
                ws.Row(x).Cell("G").Value = double.Parse(notas[i]?.Valor ?? "0");
                ws.Row(x).Cell("H").Value = double.Parse(notas[i]?.Aliquota ?? "0");
                ws.Row(x).Cell("I").FormulaA1 = $"=(G{x} / 100) * H{x}";
                ws.Row(x).Cell("J").Value = Path.GetFileNameWithoutExtension(notas[i].ChaveAcesso);
                x += 1;
                
            }
            ws.Columns().AdjustToContents();
            string path = $"RelatórioProcessado_{DateTime.Now.ToString("dd-MM-yyyy hh-mm-ss")}.xlsx";
            wb.SaveAs(path);
            return path;
        }
    }
    static void GatherMunicipio()
    {
        string munFile = "municipios.json";
        string estFile = "estados.json";
        string munJsonFile = File.ReadAllText(munFile);
        string estJsonFile = File.ReadAllText(estFile);

        List<Municipio> munJson = JsonSerializer.Deserialize<List<Municipio>>(munJsonFile) ?? new List<Municipio>();
        List<Estado> estJson = JsonSerializer.Deserialize<List<Estado>>(estJsonFile) ?? new List<Estado>();
        if(munJson.Count == 0 || estJson.Count == 0) throw new Exception("Ocorreu um erro ao converter o arquivo JSON.");
        foreach(Municipio mun in munJson)
        {
            municipios.Add(new Municipio
            {
                codigo_ibge = mun.codigo_ibge,
                nome = $"{mun.nome} - {estJson.FirstOrDefault(x => x.codigo_uf == mun.codigo_uf)?.uf ?? throw new Exception("UF não encontrada")}"
            });
        }
    }
    static async Task<bool> VerifyStatus(string xml, HttpClient client)
    {
        int[] eventos =
        {
            101101,
            101103,
            105102,
            105104,
            105105
        };
        try
        {
            foreach(int evento in eventos)
            {
                string url = $"https://sefin.nfse.gov.br/SefinNacional/nfse/{Path.GetFileNameWithoutExtension(xml)}/eventos/{evento}/1";
                HttpResponseMessage resp = await client.GetAsync(url);
                Console.WriteLine(resp.StatusCode);
                if(resp.StatusCode == HttpStatusCode.OK)
                    return true;
                else if(resp.StatusCode == HttpStatusCode.InternalServerError) throw new Exception("Ocorreu um erro com o servidor, tente novamente mais tarde.");
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex);
            throw;
        }
        return false;
    }
}
