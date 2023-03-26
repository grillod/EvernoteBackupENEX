using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Web;
using System.Security.Cryptography;

namespace EverNoteArchive
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        const string EnexFilePath = "D:\\tmp\\";
        const string NoteBookIdentifierPrefix = "ZZ_";
        const string RootOutputFolder = "D:\\MailBackups\\EverNoteBackup";
        const string HTML_TEMPLATE = "<!DOCTYPE html><html><body>{0}<br><br><u>Attachments:</u><br>{1}</body></html>";
        const string ATTACHMENT_TEMPLATE = "<a href=\"{0}\"> {1} </a><br>";
      
        

        private void button1_Click(object sender, EventArgs e)
        {
            
            
            label1.Text = "Processing ENEX Files";
            Doit();
            label1.Text = "Cleaning up ENEX Files";
            CleanUpEnexInputFiles();
            MessageBox.Show("Done");

        }

       private void  CleanUpEnexInputFiles()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(EnexFilePath);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        private void Doit()
        {

            var EnexFiles = Directory.EnumerateFiles(EnexFilePath, "*.enex");

            foreach (string currentFile in EnexFiles)
            {
                XmlDocument doc = new XmlDocument();
                doc.PreserveWhitespace = true;
                doc.Load(currentFile);

                using (XmlNodeList elemList = doc.GetElementsByTagName("note"))
                {
                    LoopThroughAllNotes(elemList);
                }
            }
        }

        private void LoopThroughAllNotes(XmlNodeList noteNodes)
        {
            for (int i = 0; i < noteNodes.Count; i++)
            {
                HandleNote(noteNodes[i]);
            }
        }

        private void HandleNote(XmlNode noteNode)
        {
            string FolderName = "";
            string Title = "";
            string NoteBody = GetNoteBody(noteNode);

            GetNoteInfo(noteNode, out Title, out FolderName);

            string NotePath = CreateNoteFolderIfNeeded(FolderName, Title);
            DeleteOnlyFilesUnderTheNoteFolder(NotePath);

            string AttachmentSection = ExtractFiles(noteNode, NotePath);

            CreateMainNoteHTMLFile(Title, NotePath, NoteBody, AttachmentSection);
        }

        private void GetNoteInfo(XmlNode noteNode, out string Title, out string FolderName)
        {
            Title = noteNode["title"].InnerText;
            FolderName = "";

            using (XmlNodeList Tags = noteNode.SelectNodes("tag"))
            {
                FolderName = GetNotebookFolderName(Tags);
            }

        }

        private string CreateNoteFolderIfNeeded(string folderName, string title)
        {
            string Template = "{0}\\{1}\\{2}-{3}";
            folderName = NormalizeFolderAndFileNames(folderName);
            title = NormalizeFolderAndFileNames(title);
            //Template = String.Format(Template, RootOutputFolder, folderName, title,Guid.NewGuid().ToString());
            Template = String.Format(Template, RootOutputFolder, folderName, title, CreateHashString(title));
            Directory.CreateDirectory(Template);
            return Template;

        }

        private void DeleteOnlyFilesUnderTheNoteFolder(string NotePath)
        {
            if (NotePath.Contains(RootOutputFolder))
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(NotePath);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
            }
        }

        private string ExtractFiles(XmlNode noteNode, string outputPath)
        {
            string AttachmentSection = "";

            XmlNodeList Resources = noteNode.SelectNodes("resource");

            foreach (XmlNode resource in Resources)
            {
                XmlNode data = resource["data"];
                XmlNode resourceAttributes = resource["resource-attributes"];
                string defaultExtension = GetDefaultExtension(resource);

                string fileName = "";
                string relativeFileName = "";
                string diskPath = "";

                if (ThereAreNoResourceAttributes(resourceAttributes))
                {
                    HandleNullResources(noteNode, out diskPath, out fileName, defaultExtension, outputPath);
                }
                else
                {
                    HandleResources(resourceAttributes, out diskPath, out fileName, defaultExtension, outputPath);
                }

                if (diskPath == "") return "";

                EnsureUniqueFilePath(ref diskPath, ref fileName);

                File.WriteAllBytes(diskPath, Convert.FromBase64String(data.InnerText));
                relativeFileName = ".\\" + fileName;
                string Link = String.Format(ATTACHMENT_TEMPLATE, relativeFileName, Path.GetFileName(fileName));
                AttachmentSection += Link;
            }

            return AttachmentSection;
        }

        string NormalizeFolderAndFileNames(string fileName)
        {
            string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

            foreach (char c in invalid)
            {
                fileName = fileName.Replace(c.ToString(), "");
            }

            return fileName;
        }

        
        private void CreateMainNoteHTMLFile(string Title, string NotePath,string NoteBody, string AttachmentSection)
        {
            NoteBody = String.Format(HTML_TEMPLATE, NoteBody, AttachmentSection);
            Title = NormalizeFolderAndFileNames(Title);

            if (Title == "") Title = "index";

            NotePath += "\\" + Title.Trim(' ') + ".html";

            File.WriteAllText(NotePath, NoteBody, System.Text.Encoding.UTF8);
        }

        string GetDefaultExtension(XmlNode resource)
        {
            string defaultExtension = "png";

            XmlNode mimeType = resource["mime"];
            if (mimeType != null)
            {
                int index = mimeType.InnerXml.IndexOf('/');
                defaultExtension = mimeType.InnerXml.Substring(index + 1);
            }

            return defaultExtension;
        }

        private bool ThereAreNoResourceAttributes(XmlNode resourceAttributes)
        {
            if (resourceAttributes == null || resourceAttributes["file-name"] == null) return true;
            return false;
        }


        private void HandleNullResources(XmlNode noteNode, out string diskPath, out string fileName, string defaultExtension, string outputPath)
        {
            XmlNode content = noteNode["content"];
            string strBody = content.InnerText;
            diskPath = "";
            fileName = "";

            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;


            strBody = strBody.Replace('&', ' ');
            doc.LoadXml(strBody);

            using (XmlNodeList elemList = doc.GetElementsByTagName("en-media"))
            {
                XmlNode mediaNode = elemList[0];
                if (mediaNode == null) return;

                string mediaType = mediaNode.Attributes["type"].Value;
                int index1 = mediaType.IndexOf('/');

                string extension = mediaType.Substring(index1 + 1);
                if (extension == "") extension = defaultExtension;

                fileName = Guid.NewGuid().ToString() + "." + extension;
                diskPath = outputPath + "\\" + fileName;
            }
        }

        private void HandleResources(XmlNode resourceAttributes, out string diskPath, out string fileName, string defaultExtension, string outputPath)
        {
            fileName = resourceAttributes["file-name"].InnerText;
            fileName = NormalizeFolderAndFileNames(fileName);

            if (fileName == "")
            {
                fileName = Guid.NewGuid().ToString() + "." + defaultExtension;
            }

            diskPath = outputPath + "\\" + fileName;
        }

        private void EnsureUniqueFilePath(ref string diskPath, ref string fileName)
        {
            if (File.Exists(diskPath))
            {
                string dirName = Path.GetDirectoryName(diskPath) + "\\";
                string newFileName = Guid.NewGuid().ToString() + Path.GetExtension(diskPath);
                fileName = newFileName;
                dirName += newFileName;
                diskPath = dirName;
            }
        }

        private string GetNotebookFolderName(XmlNodeList Tags)
        {
            
            foreach(XmlNode node in Tags)
            {
                string innerText = node.InnerText;

                if (innerText.Length < NoteBookIdentifierPrefix.Length) continue;

                if (innerText.Substring(0, NoteBookIdentifierPrefix.Length) == NoteBookIdentifierPrefix)
                {
                    return innerText.Substring(NoteBookIdentifierPrefix.Length);
                }
            }

            return "Unfiled";
        }

        private string GetNoteBody(XmlNode noteNode)
        {
            XmlNode ContentNode = noteNode["content"];
            string strBody = ContentNode.InnerText;
            return strBody;
        }


        private string CreateHashString(string seed)
        {
            using (var sha = new System.Security.Cryptography.SHA256Managed())
            {
                // Convert the string to a byte array first, to be processed
                byte[] textBytes = System.Text.Encoding.UTF8.GetBytes(seed);
                byte[] hashBytes = sha.ComputeHash(textBytes);

                // Convert back to a string, removing the '-' that BitConverter adds
                string hash = BitConverter
                    .ToString(hashBytes)
                    .Replace("-", String.Empty);

                hash = hash.Substring(0, 15);

                return hash;
            }
        }
    }
}
