using System;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ConsoleApp
{
    class Program
    {
        #region Properties

        private static String ImagePath
        {
            get { return Environment.CurrentDirectory + @"\Images\docx-icon.png"; }
        }

        #endregion

        #region Methods

        static void Main(string[] args)
        {
            Console.Title = "Docx tests";
            ShowMenu();
        }

        private static void ShowMenu()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Menu");
                Console.WriteLine("\n1. Hello word;");
                Console.WriteLine("2. Word documento with table header and footer;");
                Console.WriteLine("0. Exit");
                Console.Write("\nChoose a option: ");

                String op = Console.ReadLine().Trim();

                switch (op)
                {
                    case "1":
                        CreateHelloWorldDoc();
                        break;

                    case "2":
                        CreateDocWithTable();
                        break;

                    case "0":
                        Console.WriteLine("Type any key to exit.");
                        Console.ReadKey();
                        Environment.Exit(0);
                        break;

                    default:
                        Console.WriteLine("Choose a correct option from menu.");
                        Console.ReadKey();
                        break;
                }
            }
        }

        private static void CreateHelloWorldDoc()
        {
            String path = FileDirectoryHelper.GetDirPath(ConfigurationManager.AppSettings["SAMPLE_FILES_DIR_PATH"]) + @"\doc-sample.docx";
            DocX doc = DocX.Create(path);

            Formatting docHeaderFormat = new Formatting();
            docHeaderFormat.FontFamily = new Font("Calibri");
            docHeaderFormat.Size = 10.0;
            docHeaderFormat.Bold = true;

            Formatting docBodyFormat = new Formatting();
            docBodyFormat.FontFamily = new Font("Calibri");
            docBodyFormat.Size = 12.0;

            Formatting docFooterFormat = new Formatting();
            docFooterFormat.FontFamily = new Font("Calibri");
            docFooterFormat.Size = 10.0;
            docFooterFormat.Bold = true;

            Formatting docTitleFormat = new Formatting();
            docTitleFormat.FontFamily = new Font("Calibri");
            docTitleFormat.Size = 16.0;

            #region Header

            doc.AddHeaders();
            Header header = doc.Headers.Odd;
            Paragraph parHeader = header.InsertParagraph();
            Image docxIcon = doc.AddImage(ImagePath);
            Picture picture = docxIcon.CreatePicture(25, 25);
            parHeader.AppendPicture(picture);
            parHeader.Append(" Powered by Docx api", docHeaderFormat);

            #endregion

            #region Body

            doc.InsertParagraph("Hello world (a nonsense text)", false, docTitleFormat).Alignment = Alignment.center;

            StringBuilder fullText = new StringBuilder();
            Random a = new Random();
            Random b = new Random();

            for (Int32 i = 0; i < 1000; i++)
            {
                Int32 wordSize = b.Next(1, 15);
                StringBuilder word = new StringBuilder();

                for (Int32 j = 0; j < wordSize; j++)
                {
                    String chara = Char.ConvertFromUtf32(a.Next(32, 126));

                    if (Char.IsLetter(Convert.ToChar(chara)) && ! Char.IsWhiteSpace(Convert.ToChar(chara)))
                    {
                        word.Append(j == 0 ? chara.ToUpper() : chara.ToLower());
                    }
                }

                word.Append((Char)32);
                fullText.Append(word);
            }

            doc.InsertParagraph(fullText.ToString(), false, docBodyFormat).Alignment = Alignment.both;

            #endregion

            #region Footer

            doc.AddFooters();

            Footer footer = doc.Footers.Odd;
            Paragraph parFooter =
                footer.InsertParagraph
                (
                    String.Format("Document created in {0} as {1}", DateTime.Now.ToString("dd/MM/yyyy"), DateTime.Now.ToString("HH:mm:ss"))
                    , false
                    , docFooterFormat
                );

            #endregion

            SaveAndOpen(doc, path);
        }

        private static void CreateDocWithTable()
        {
            String path = FileDirectoryHelper.GetDirPath(ConfigurationManager.AppSettings["SAMPLE_FILES_DIR_PATH"]) + @"\doc-with-table-sample.docx";
            DocX doc = DocX.Create(path);

            Formatting docHeaderFormat = new Formatting();
            docHeaderFormat.FontFamily = new Font("Calibri");
            docHeaderFormat.Size = 10.0;
            docHeaderFormat.Bold = true;

            Formatting docTitleFormat = new Formatting();
            docTitleFormat.FontFamily = new Font("Calibri");
            docTitleFormat.Size = 16.0;

            Formatting docFooterFormat = new Formatting();
            docFooterFormat.FontFamily = new Font("Calibri");
            docFooterFormat.Size = 10.0;
            docFooterFormat.Bold = true;

            #region Header

            doc.AddHeaders();
            Header header = doc.Headers.Odd;
            Paragraph parHeader = header.InsertParagraph();
            Image docxIcon = doc.AddImage(ImagePath);
            Picture picture = docxIcon.CreatePicture(25, 25);
            parHeader.AppendPicture(picture);
            parHeader.Append(" Powered by Docx api", docHeaderFormat);

            #endregion

            #region Body

            Paragraph paragraph = doc.InsertParagraph("Document with embeded table below", false, docTitleFormat);
            paragraph.Alignment = Alignment.center;
            paragraph.AppendLine();

            Int32 numColumns;
            Int32 numRows;
            Int32 minColumns = 1;
            Int32 maxColumns = 4;
            Int32 minRows = 1;
            Int32 maxRows = 40;

            while (true)
            {
                Console.Write("Type a number of columns: ");
                String input = Console.ReadLine();

                if (String.IsNullOrEmpty(input))
                {
                    Console.WriteLine("No input typed. Assuming max value {0} to columns.", maxColumns);
                    numColumns = maxColumns;
                    Console.ReadKey();
                    break;
                }

                numColumns = Convert.ToInt32(input);

                if (numColumns < minColumns || numColumns > maxColumns)
                {
                    Console.WriteLine("Type a number between {0} and {1}", minColumns, maxColumns);
                    Console.ReadKey();
                    continue;
                }
                break;
            }

            while (true)
            {
                Console.Write("Type a number of rows: ");
                String input = Console.ReadLine();

                if (String.IsNullOrEmpty(input))
                {
                    Console.WriteLine("No input typed. Assuming max value {0} to rows.", maxRows);
                    numRows = maxRows;
                    Console.ReadKey();
                    break;
                }

                numRows = Convert.ToInt32(input);

                if (numRows < 0 || numRows > maxRows)
                {
                    Console.WriteLine("Type a number between {0} and {1}", minRows, maxRows);
                    Console.ReadKey();
                    continue;
                }
                break;
            }

            Table table = doc.AddTable(numRows, numColumns);
            table.Alignment = Alignment.center;
            table.Design = TableDesign.ColorfulList;
            table.AutoFit = AutoFit.Contents;

            #region Update content cells

            for (Int32 r = 0; r < numRows; r++)
            {
                for (Int32 c = 0; c < numColumns; c++)
                {
                    table.Rows[r].Cells[c].Paragraphs.First().Alignment = Alignment.center;
                    table.Rows[r].Cells[c].Paragraphs.First().Append(String.Format("{0} {1} - Column {2}", (r == 0 ? "Header" : "Row"), r + (r == 0 ? 1 : 0), c + 1));
                }
            }

            #endregion

            #endregion

            #region Footer

            doc.AddFooters();

            Footer footer = doc.Footers.Odd;
            Paragraph parFooter =
                footer.InsertParagraph
                (
                    String.Format("Document created in {0} as {1}", DateTime.Now.ToString("dd/MM/yyyy"), DateTime.Now.ToString("HH:mm:ss"))
                    , false
                    , docFooterFormat
                );

            #endregion

            doc.InsertTable(table);
            SaveAndOpen(doc, path);
        }

        private static void SaveAndOpen(DocX doc, String path)
        {
            doc.Save();

            String pathProgram = GetPathProgram();

            if (String.IsNullOrEmpty(pathProgram))
            {
                Console.WriteLine("Without program to open.");
                return;
            }

            Process.Start(pathProgram, path);
        }

        private static String GetPathProgram()
        {
            String pathProgram = String.Empty;

            if (ConfigurationManager.AppSettings["FULL_PATH_PROGRAM"] != null)
            {
                if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["FULL_PATH_PROGRAM"]))
                {
                    pathProgram = ConfigurationManager.AppSettings["FULL_PATH_PROGRAM"];
                }
            }

            return pathProgram;
        }

        #endregion
    }
}
