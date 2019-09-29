using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ionic.Zip;
using Syncfusion.Pdf;
using System.IO;
using Spire.Pdf;
using Spire.Doc;
using SautinSoft.Document;

namespace Problem1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public bool IsBackground { get; private set; }

        private void btnFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select path";
            if (fbd.ShowDialog() == DialogResult.OK)
                txtFolder.Text = fbd.SelectedPath;

        }

        private void btnZipFolder_Click(object sender, EventArgs e)
        {

            //try4 WORKS
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text) ||
                string.IsNullOrEmpty(textBox3.Text))
                MessageBox.Show("Please select file 1, file 2 and file 3");
            else
            {
                using (ZipFile zip = new ZipFile())
                {
                    // add this map file into the "images" directory in the zip archive
                    zip.AddFile(textBox1.Text);
                    // add the report into a different directory in the archive
                    zip.AddFile(textBox2.Text);
                    zip.AddFile(textBox3.Text);
                    zip.Save("C://Users//dell//Desktop//Input//outputproblemstatement1.7z.zip");
                }
            }

            ////try1

            //string startPath = txtFolder.Text;
            //string zipPath = txtFolder.Text;
            //ZipFile.CreateFromDirectory(startPath, zipPath);


            // //try2

            // if (string.IsNullOrEmpty(txtFolder.Text))
            // {
            //     MessageBox.Show("Select folder", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //     txtFolder.Focus();
            //     return;

            // }
            // string path = txtFolder.Text;
            // Thread thread = new Thread(t =>
            // {
            //     using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
            //     {

            //         //zip.UseUnicodeAsNecessary = true;
            //         zip.AddDirectory(path);
            //         //zip.AddFile(path,path);
            //         System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(path);
            //         zip.SaveProgress += Zip_SaveProgress;
            //         zip.Save(string.Format("{0}{1}.zip", di.Parent.FullName, di.Name));
            //         //zip.Save(string.Format("outputproblemstatement1.zip"));
            //     }
            // }
            //);
            // { IsBackground = true; };
            // thread.Start();

            ////try3

            //using (ZipFile zip = new ZipFile())
            //{
            //    var path = txtFolder.Text;
            //    zip.UseUnicodeAsNecessary = true;  // utf-8
            //    zip.AddDirectory(@txtFolder.Text);
            //    zip.Comment = "This zip was created at " + System.DateTime.Now.ToString("G");
            //    zip.Save(path);
            //}



        }

        private void Zip_SaveProgress(object sender, SaveProgressEventArgs e)
        {
            if (e.EventType == Ionic.Zip.ZipProgressEventType.Saving_BeforeWriteEntry)
            {
                progressBar1.Invoke(new MethodInvoker(delegate

            {
                progressBar1.Maximum = e.EntriesTotal;
                progressBar1.Value = e.EntriesSaved + 1;
            }));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //    //try1

            //    //Creates a PDF document
            //    Spire.Pdf.PdfDocument finalDoc = new Spire.Pdf.PdfDocument();

            //    //Creates a string array of source files to be merged
            //    string[] source = { textBox1.Text, textBox2.Text, textBox3.Text };

            //    //Merge PDF documents
            //    Spire.Pdf.PdfDocumentBase.Merge(finalDoc, source);
            //    //Save the document
            //    finalDoc.SaveToFile("Sample.pdf");
            //


            //try2
            
            //Spire.Pdf.PdfDocument[] documents = new Spire.Pdf.PdfDocument[4];
            //using (MemoryStream ms1 = new MemoryStream())
            //{
            //    Document doc = new Document("c://users//dell//desktop//input//second.docx", Spire.Doc.FileFormat.Auto);
            //    doc.SaveToStream(ms1, Spire.Doc.FileFormat.PDF);
            //    documents[0] = new Spire.Pdf.PdfDocument(ms1);
            //}

            //using (MemoryStream ms2 = new MemoryStream())
            //{
            //    Document jpg = new Document("c://users//dell//desktop//input//first.jpg", Spire.Doc.FileFormat.Auto);
            //    jpg.SaveToStream(ms2, Spire.Doc.FileFormat.PDF);
            //    documents[1] = new Spire.Pdf.PdfDocument(ms2);
            //}
            //using (MemoryStream ms3 = new MemoryStream())
            //{
            //    Document pdf = new Document("c://users//dell//desktop//input//third.pdf", Spire.Doc.FileFormat.Auto);
            //    pdf.SaveToStream(ms3, Spire.Doc.FileFormat.PDF);
            //    documents[2] = new Spire.Pdf.PdfDocument(ms3);

            //}
            //documents[3] = new Spire.Pdf.PdfDocument("fourth.pdf");
            //for (int i = 2; i > -1; i--)
            //{
            //    documents[3].AppendPage(documents[i]);
            //}

            //documents[3].SaveToFile("outputproblemstatement2.pdf");


            //try3

            // Path to our combined document.
            string singlePDFPath = "Single.pdf";
            string workingDir = Path.GetFullPath(@"C:\Users\dell\Desktop\Input");

            List<string> supportedFiles = new List<string>();
            foreach (string file in Directory.GetFiles(workingDir, "*.*"))
            {
                string ext = Path.GetExtension(file).ToLower();

                if (ext == ".docx" || ext == ".pdf" || ext == ".txt")
                    supportedFiles.Add(file);
            }

            // Create single pdf.
            DocumentCore singlePDF = new DocumentCore();

            foreach (string file in supportedFiles)
            {
                DocumentCore dc = DocumentCore.Load(file);

                Console.WriteLine("Adding: {0}...", Path.GetFileName(file));

                // Create import session.
                ImportSession session = new ImportSession(dc, singlePDF, StyleImportingMode.KeepSourceFormatting);

                // Loop through all sections in the source document.
                foreach (SautinSoft.Document.Section sourceSection in dc.Sections)
                {
                    
                    SautinSoft.Document.Section importedSection = singlePDF.Import<SautinSoft.Document.Section>(sourceSection, true, session);    
                    if (dc.Sections.IndexOf(sourceSection) == 0)
                        importedSection.PageSetup.SectionStart = SectionStart.NewPage;
                    singlePDF.Sections.Add(importedSection);
                }
            }

            
            singlePDF.Save(singlePDFPath);

            ////try4
            //string[] inputFilePaths = Directory.GetFiles(txtFolder.Text, "first.jpg");
            //Console.WriteLine("Number of files: {0}.", inputFilePaths.Length);
            //using (var outputStream = File.Create(txtFolder.Text))
            //{
            //    foreach (var inputFilePath in inputFilePaths)
            //    {
            //        using (var inputStream = File.OpenRead(inputFilePath))
            //        {
            //            // Buffer size can be passed as the second argument.
            //            inputStream.CopyTo(outputStream);
            //        }
            //        Console.WriteLine("The file {0} has been processed.", inputFilePath);
            //    }
            //}

        } 

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFolder.Text) || string.IsNullOrEmpty(txtdestinatie.Text))
            {
                MessageBox.Show("Please select Folder and Destination");
            }
            else { 
                //Create e directories
                foreach (string dirPath in Directory.GetDirectories(txtFolder.Text, "*",
                SearchOption.AllDirectories))
                    Directory.CreateDirectory(dirPath.Replace(txtFolder.Text, txtdestinatie.Text));

                //Copy all the files 
                foreach (string newPath in Directory.GetFiles(txtFolder.Text, "*.*",
                    SearchOption.AllDirectories))
                    File.Copy(newPath, newPath.Replace(txtFolder.Text, txtdestinatie.Text), true);
                MessageBox.Show("Files were copied to " + txtdestinatie.Text);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "All Files|*.*", ValidateNames = true, Multiselect = false })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                    textBox1.Text = ofd.FileName;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "All Files|*.*", ValidateNames = true, Multiselect = false })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                    textBox2.Text = ofd.FileName;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "All Files|*.*", ValidateNames = true, Multiselect = false })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                    textBox3.Text = ofd.FileName;
            }
        }

        private void btndestinatie_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select path";
            if (fbd.ShowDialog() == DialogResult.OK)
                txtdestinatie.Text = fbd.SelectedPath;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
        
        }

        private void button2_Click_2(object sender, EventArgs e)
        {

        }
    }
    }
     
