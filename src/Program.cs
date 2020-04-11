//////////////////////////////////////
////// Springer Book Downloader //////
////// Written by yak112 - 2020 //////
//////////////////////////////////////
using System;
using System.Net;
using System.ComponentModel;
using System.Threading;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Springer_webscrap
{
   class Downloader
    {
        private volatile bool _complete;
        private string DownloadFile(string pathToDownload, string filename, string documentURL)
        {
            string filePath = pathToDownload + "\\" +filename;
            var di = new DirectoryInfo(pathToDownload);
            if(!di.Exists)
            {
                di.Create();
            }

            //Download the book list in a temporal file
            WebClient springerWeb = new WebClient();
            var link = new Uri(documentURL);
            _complete = false;

            springerWeb.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
            springerWeb.DownloadProgressChanged += new DownloadProgressChangedEventHandler(DownloadProgress);
            springerWeb.DownloadFileAsync(link, filePath);
            while (springerWeb.IsBusy) { Thread.Sleep(1000); }

            return filePath;
        }
        public int ParseExcelAndDownload(string pathToDownload) {
            string bookListURL = "https://resource-cms.springernature.com/springer-cms/rest/v1/content/17858272/data/v4";
            string bookCategory = "";
            string bookAuthor = "";
            string bookTitle = "";
            int row_num = 0;
            Console.WriteLine("Starting book list download...");
            string tempFile = DownloadFile(pathToDownload,"booklist.tmp", bookListURL);

            Console.WriteLine("Starting list parsing and book download...");
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(tempFile, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                foreach (Row r in sheetData.Elements<Row>())
                {
                    if (row_num > 0)
                    {
                        int field_num = 1;
                        foreach (Cell c in r.Elements<Cell>())
                        {
                            if (c.DataType != null && c.DataType == CellValues.SharedString)
                            {
                                var stringId = Convert.ToInt32(c.InnerText);
                                if (field_num == 12) //Create a folder based on english package name
                                {
                                    bookCategory = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                    string pathString = System.IO.Path.Combine(pathToDownload, bookCategory);
                                    System.IO.Directory.CreateDirectory(pathString);
                                }
                                if (field_num == 2) //Gets book authors
                                {
                                    bookAuthor = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                }
                                if (field_num == 1) //Gets book title
                                {
                                    bookTitle = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                }
                                if (field_num == 19) ////Gets book title
                                {
                                    //PDF Book
                                    string bookLink = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText;
                                    string filename = "";
                                    string filePath = "";

                                    //There are some redirects in Springer webpage, create a HttpWebRequest to handle this redirects
                                    HttpWebRequest springerRequest = (HttpWebRequest)WebRequest.Create(bookLink);
                                    springerRequest.MaximumAutomaticRedirections = 3;
                                    springerRequest.AllowAutoRedirect = true;
                                    HttpWebResponse springerResponse = (HttpWebResponse)springerRequest.GetResponse();
                                    bookLink = springerResponse.ResponseUri.ToString();
                                    bookLink = bookLink.Replace("/book/", "/content/pdf/");
                                    bookLink = bookLink.Replace("%2F", "/");
                                    bookLink += ".pdf";
                                    Console.WriteLine($"Downloading book {bookTitle}.");
                                    bookAuthor = bookAuthor.Replace(", ", "-").Replace(".", "").Replace("/", " ");
                                    bookTitle = bookTitle.Replace(", ", "-").Replace(".", "").Replace("/", " ");
                                    filePath = pathToDownload + "\\" + bookCategory + "\\";
                                    filename = bookAuthor + "\\" + bookTitle + ".pdf";
                                    string pathString = System.IO.Path.Combine(pathToDownload + "\\" + bookCategory, bookAuthor);
                                    System.IO.Directory.CreateDirectory(pathString);

                                    DownloadFile(filePath, filename, bookLink);

                                    //Epub book
                                    bookLink = springerResponse.ResponseUri.ToString();
                                    bookLink = bookLink.Replace("/book/", "/download/epub/");
                                    bookLink = bookLink.Replace("%2F", "/");
                                    bookLink += ".epub";
                                    filePath = pathToDownload + "\\" + bookCategory + "\\";
                                    filename = bookAuthor + "\\" + bookTitle + ".epub";

                                    DownloadFile(filePath, filename, bookLink);
                                }
                            }
                            field_num++;
                        }
                    }
                    row_num++;
                }
            }
            //Time to do some cleanup
            Console.WriteLine("Doing some cleanup. This may take a few moments, please wait...");
            var di = new DirectoryInfo(pathToDownload);
            File.Delete(tempFile);
            FileInfo[] zeroSizeFiles = di.GetFiles("*.*",SearchOption.AllDirectories).Where(fi => fi.Length == 0).ToArray();
            foreach (FileInfo file in zeroSizeFiles) {
                file.Delete();
            }

            return row_num;
        }
        public bool DownloadCompleted { get { return _complete; } }

        private void DownloadProgress(object sender, DownloadProgressChangedEventArgs e)
        {
            // Displays the operation identifier, and the transfer progress.
            Console.WriteLine("{0}    downloaded {1} of {2} bytes. {3} % complete...",(string)e.UserState,e.BytesReceived,e.TotalBytesToReceive,e.ProgressPercentage);
        }

        private void Completed(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                Console.WriteLine("Download has been canceled.");
            }
            else
            {
                Console.WriteLine("Download finished!");
            }

            _complete = true;
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Downloader FileDownloader = new Downloader();

            Console.Title = "Springer Textbook Downloader";
            Console.WriteLine("----------------------------");
            Console.WriteLine("Springer Textbook Downloader");
            Console.WriteLine("Coded by yak112 - 2020");
            Console.WriteLine("----------------------------");
            Console.WriteLine("Please type in the folder where you want to store the books:");
            string dest_path = Console.ReadLine();
            Console.WriteLine("Working on it...");
            int result = FileDownloader.ParseExcelAndDownload(dest_path);
            while (!FileDownloader.DownloadCompleted)
                Thread.Sleep(1000);
            if (result > 0)
            {
                Console.WriteLine("\n\n----------------------------");
                Console.WriteLine($"Downloaded {result} files.");
                Console.WriteLine("Work has finished. Enjoy the books!"); 
                Console.WriteLine("Thanks Springer!!");
                Console.ReadKey();
            } else
            {
                Console.WriteLine("Something has gone bad, please retry again later.");
                Console.ReadKey();
            }
        }
    }
}
