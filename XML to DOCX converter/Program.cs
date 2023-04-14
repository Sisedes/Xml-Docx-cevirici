using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;
class XMLtoDOCX
{
    public void XmlToDocx(string xmlFilePath, string docxFilePath)
    {
        XmlDocument xml = new XmlDocument();
        xml.Load(xmlFilePath); //xml ogesi olusturma ve metoda gelen dosya konumundan yukleme

        
        using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document))//belirtilen konumdaki docx dosyasının üzerine yazıyor
        {
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();
            foreach (XmlNode node in xml.DocumentElement.ChildNodes)
            {
                Paragraph paragraph = new Paragraph();
                Run run = new Run();
                run.AppendChild(new Text(node.InnerText));
                paragraph.Append(run);
                body.Append(paragraph);
            }
            mainPart.Document.Append(body);
        }


    }

    public string DocxOlustur(string kaynakdosyaAD)
    {
        string filePath = @"C:\Users\kaan4\source\repos\XML to DOCX converter\kaynak\"+kaynakdosyaAD+".docx";
        using (WordprocessingDocument document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
        {
            
            MainDocumentPart mainPart = document.AddMainDocumentPart(); //yeni docx olusturuyor.
            mainPart.Document = new Document();
            mainPart.Document.Save();//kaydediyor.
        }

        return filePath;
    }

    static void Main(string[] args)
    {
        XMLtoDOCX nesne = new XMLtoDOCX();
        #region hazir docx uzerine yazmak icin

        string filepathofXML = "C:\\Users\\kaan4\\source\\repos\\XML to DOCX converter\\kaynak\\books.xml";
        string filepathofDOCX = "C:\\Users\\kaan4\\source\\repos\\XML to DOCX converter\\kaynak\\deneme.docx"; //burada hazır verilen docx dosyası üzerine **yazma** işlemi yapılıyor

        nesne.XmlToDocx(filepathofXML,filepathofDOCX);

        #endregion

        #region yeni bir docx uzerine yazmak icin
        //xml filepathi yine belirtmek zorundayız
        //yeniden olusturmayi tercih ettim.
        string filepathofXML2 = "C:\\Users\\kaan4\\source\\repos\\XML to DOCX converter\\kaynak\\books.xml";
        string yenidocxPath =nesne.DocxOlustur("yenidosya");
        nesne.XmlToDocx(filepathofXML2,yenidocxPath);
        #endregion

    }
}
