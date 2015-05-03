using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Linq;

namespace d2xml
{
    public class Builder
    {
        public Builder ()
        {

        }
        
        public bool getDocumentsToXml()
        {
                string outputFile = ConfigurationManager.AppSettings["outputFile"].ToString();
                bool lowercaseFirstLetter = bool.Parse(ConfigurationManager.AppSettings["lowercaseFirstLetter"].ToString());

                FileInfo fi = new FileInfo(outputFile);
                if (fi.Exists)
                {
                    fi.Delete();
                }

                string[] fileList = getFileList();

                List<List<string>> vocab = new List<List<string>>();

                XmlTextWriter wri = new XmlTextWriter(outputFile, null);
                wri.WriteStartDocument();
                wri.WriteStartElement("vocab");
                foreach (string file in fileList)
                {
                    List<List<string>> vocabList = new List<List<string>>();

                    string filename = Path.GetFileNameWithoutExtension(file);
                    string dateUploaded = getDate(file);

                    vocabList = getVocab(lowercaseFirstLetter, file);
                    vocab.AddRange(vocabList);

                    wri.WriteStartElement("lesson");
                    wri.WriteAttributeString("filename", filename);
                    wri.WriteAttributeString("date", dateUploaded);

                    appendXML(wri, vocabList);

                    wri.WriteEndElement();
                }
                wri.WriteEndElement();

                wri.WriteEndDocument();
                wri.Close();

                return true;    
        }

        private void appendXML(XmlTextWriter wri, List<List<string>> vocabList)
        {
            foreach (List<string> vocab in vocabList)
            {
                wri.WriteStartElement("line");

                wri.WriteStartElement("english");
                wri.WriteString (vocab[1]);
                wri.WriteEndElement();

                wri.WriteStartElement("language");
                wri.WriteString(vocab[0]);
                wri.WriteEndElement();

                wri.WriteEndElement();
            }
        }

        private List<List<string>> getVocab(bool lcFirstLetter, string file)
        {
            List<List<string>> vocab = new List<List<string>>();

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(file, false))
                {
                    
                    Body body = doc.MainDocumentPart.Document.Body;
                    Table table = body.ChildElements.OfType<Table>().ToArray<Table>() [0];
                                     
                    TableRow [] rows = table.ChildElements.OfType<TableRow>().ToArray();
                    foreach (TableRow row in rows)
                    {
                        string language = string.Empty;
                        string english = string.Empty;

                        TableCell [] cells = row.ChildElements.OfType<TableCell>().ToArray();
                        
                        language = cells [0].InnerText;                        
                        english = cells [1].InnerText;

                        if (language.Length > 0 && english.Length > 0)
                        {
                            if (lcFirstLetter)
                            {
                                language = language[0].ToString().ToLower() + language.Substring(1);
                                english = english[0].ToString().ToLower() + english.Substring(1);
                            }
                            List<string> vocabline = new List<string>() { language, english };

                            vocab.Add(vocabline);
                        }
                     } 
                     
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error in " + file + " " + ex.Message);
            }

            return vocab;
        }

        private string getDate(string file)
        {
            FileInfo fi = new FileInfo(file);
            DateTime date = fi.CreationTime;

            return date.ToString ("yyyy-MM-dd");

        }

        private string[] getFileList()
        {
            string basedir = ConfigurationManager.AppSettings["baseDir"].ToString();

            string[] files = getFiles(basedir);
            List<string> fileList = new List<string>();

            foreach (string file in files)
            {
                fileList.Add(file);
            }

            string[] dirs = Directory.GetDirectories(basedir);
            foreach (string dir in dirs)
            {
                string[] files2 = getFiles(dir);

                foreach (string file in files2)
                {
                    fileList.Add(file);
                }

            }

            return fileList.ToArray();
        }

        private string[] getFiles(string dir)
        {
            string [] files = Directory.GetFiles(dir, "*.docx");
            return files;
        }

    }

}
