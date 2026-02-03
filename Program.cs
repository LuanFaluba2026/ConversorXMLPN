using System.ComponentModel.DataAnnotations;
using System.Runtime.InteropServices.Marshalling;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

class NotaFiscal
{
    private bool isCancelada = false;
    public long NumeroNota { get; set; }
    public DateTime DataEmissao { get; set; }
    public string? Nome { get; set; }
    public string? Valor { get; set; }
    public string? Aliquota { get; set; }
    public bool Cancelada {get {return isCancelada;} set {isCancelada = value;}}
}
class Program
{
    static void Main(string[] args)
    {
        try
        {
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
            List<NotaFiscal> processedXML = ProcessXML(xmlFiles, typeInput);
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
    static List<NotaFiscal> ProcessXML(string[] xmls, string typeInput)
    {
        if(xmls.Length == 0) throw new Exception("Nenhum arquivo na pasta. Tente novamete...");

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

            long numNota = long.Parse(nfseInfo?.Element(nf + "nNFSe")?.Value.ToString() ?? "0");
            DateTime dataEmi = DateTime.Parse(nfseInfo?.Element(nf + "dhProc")?.Value ?? "");
            string nome;
            if(typeInput.Equals("P", StringComparison.OrdinalIgnoreCase))
                nome = tomaInfo?.Element(nf + "xNome")?.Value ?? "";
            else
                nome = emitInfo?.Element(nf + "xNome")?.Value ?? "";
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
            
            
            NotaFiscal notaObj = new NotaFiscal
            {
                NumeroNota = numNota,
                DataEmissao = dataEmi,
                Nome = nome,
                Valor = valor,
                Aliquota = aliquotaIss,
                Cancelada = canc != null
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


            ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;
            ws.Row(1).Style.Font.Bold = true;
            ws.Row(1).Cell("A").Value = "Sequencia";
            ws.Row(1).Cell("B").Value = "Data de Emissão";
            ws.Row(1).Cell("C").Value = "Número nota";
            ws.Row(1).Cell("D").Value = "Nome";
            ws.Row(1).Cell("E").Value = "Valor Bruto";
            ws.Row(1).Cell("F").Value = "Alíquota";
            ws.Row(1).Cell("G").Value = "Valor ISSQN";

            int x = 2;
            for(int i = 0; i < notas.Count(); i++)
            {
                if(x % 2 == 0) ws.Row(x).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                else if(notas[i].Cancelada) ws.Row(x).Style.Fill.BackgroundColor = XLColor.CoralRed;
                else ws.Row(x).Style.Fill.BackgroundColor = XLColor.White;
                ws.Row(x).Cell("A").Value =  x - 1;
                ws.Row(x).Cell("B").Value = notas[i].DataEmissao.ToString("dd/MM/yyyy");
                ws.Row(x).Cell("C").Value = notas[i].NumeroNota.ToString();
                ws.Row(x).Cell("D").Value = notas[i].Nome;
                ws.Row(x).Cell("E").Value = double.Parse(notas[i]?.Valor ?? "0");
                ws.Row(x).Cell("F").Value = double.Parse(notas[i]?.Aliquota ?? "0");
                ws.Row(x).Cell("G").FormulaA1 = $"=(F{x} / 100) * E{x}";
                x += 1;
                
            }
            string path = $"RelatórioProcessado_{DateTime.Now.ToString("dd-MM-yyyy hh-mm-ss")}.xlsx";
            wb.SaveAs(path);
            return path;
        }
    }
}
