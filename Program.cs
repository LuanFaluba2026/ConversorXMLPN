using System.Runtime.InteropServices.Marshalling;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;

class NotaFiscal
{
    public int NumeroNota { get; set; }
    public DateTime DataEmissao { get; set; }
    public string? Nome { get; set; }
    public string? Valor { get; set; }
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
            string[] xmlFiles = Directory.GetFiles(path);
            List<NotaFiscal> processedXML = ProcessXML(xmlFiles, typeInput);
            GenerateSheet(processedXML);
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
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

            int numNota = int.Parse(nfseInfo?.Element(nf + "nNFSe")?.Value.ToString() ?? "0");
            DateTime dataEmi = DateTime.Parse(nfseInfo?.Element(nf + "dhProc")?.Value ?? "");
            string nome;
            if(typeInput.Equals("P", StringComparison.OrdinalIgnoreCase))
                nome = tomaInfo?.Element(nf + "xNome")?.Value ?? "";
            else
                nome = emitInfo?.Element(nf + "xNome")?.Value ?? "";
            string valor = valInfo?.Element(nf + "vServ")?.Value ?? "";

            NotaFiscal notaObj = new NotaFiscal
            {
                NumeroNota = numNota,
                DataEmissao = dataEmi,
                Nome = nome,
                Valor = valor
            };
            nfs.Add(notaObj);
        }
        return nfs.OrderBy(x => x.DataEmissao).ThenBy(x => x.NumeroNota).ToList();
    }

    static void GenerateSheet(List<NotaFiscal> notas)
    {
        using(var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Relatório de Notas Processadas");
            for(int i = 0; i < notas.Count(); i++)
            {
                int x = i + 1;
                ws.Row(x).Cell("A").Value = notas[i].DataEmissao.ToString("dd/MM/yyyy");
                ws.Row(x).Cell("B").Value = notas[i].NumeroNota;
                ws.Row(x).Cell("C").Value = notas[i].Nome;
                ws.Row(x).Cell("D").Value = notas[i].Valor;
            }
            wb.SaveAs($"RelatórioProcessado_{DateTime.Now.ToString("dd-MM-yyyy hh-mm-ss")}.xlsx");
        }
    }
}
