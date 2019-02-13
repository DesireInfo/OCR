/*
Copyright(C)

This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as published by the Free Software Foundation, either version 3 of the License, or(at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License along with this program.If not, see https://www.gnu.org/licenses/.


Also add information on how to contact you by electronic and paper mail.

If your software can interact with users remotely through a computer network, you should also make sure that it provides a way for users to get its source.For example, if your program is a web application, its interface could display a "Source" link that leads users to an archive of the code.There are many ways you could offer source, and different solutions will be better for different programs; see section 13 for the specific requirements.

You should also get your employer(if you work as a programmer) or school, if any, to sign a "copyright disclaimer" for the program, if necessary.
Full Notice : https://itextpdf.com/en/how-buy/legal/agpl-gnu-affero-general-public-license

*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Security;
using System.IO;
using System.Net;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data;

namespace PDFReader
{
    public class Program:PDFReader.PDFFile 
    {
        public static ClientContext context;
        public static int countImgPdf = 0;
        public static List<PDFFile> imgPDF = new List<PDFFile>();
        public static PDFFile objPDF = new PDFFile();
        public static System.Data.DataTable table = new System.Data.DataTable();

        public static int count = 0;


        public static DirectoryInfo dir = new DirectoryInfo(@"C:\PDF Reports");
        static void Main(string[] args)
        {

            table.Columns.Add("File Name", typeof(string));
            table.Columns.Add("File URL", typeof(string));

            Console.WriteLine("Enter Site Url:");
            context = new ClientContext(Console.ReadLine());

            Console.WriteLine("Enter UserName:");
            string emailAddress = Console.ReadLine();

            Console.WriteLine("Enter Password:");
            string password = Console.ReadLine();



            SecureString sec_pass = new SecureString();
            Array.ForEach(password.ToArray(), sec_pass.AppendChar);
            sec_pass.MakeReadOnly();

            context.Credentials = new SharePointOnlineCredentials(emailAddress, sec_pass);
            //context.Credentials = new NetworkCredential(emailAddress, password);

            context.RequestTimeout = -1;

            do
            {
                countImgPdf = 0;
                imgPDF.Clear();
                table.Rows.Clear();
                Console.WriteLine("Enter Library Name:");
                getLibFiles(Console.ReadLine());
                foreach (var item in imgPDF)
                {
                    Console.WriteLine(item.fileName);
                    Console.WriteLine(item.fileUrl);
                }
                Console.WriteLine("Pdf file with Image:- " + countImgPdf);
                if (countImgPdf != 0)
                {
                    ExportFilesToExcel();
                }
                Console.WriteLine("Press Esc to exit");
            } while (Console.ReadKey().Key != ConsoleKey.Escape);

            //createPdfFiles();

        }
        public static void getLibFiles(string libName)
        {
            try
            {
                
                Web web = context.Web;
                Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle(libName);
                Console.WriteLine("Library Name:- " + libName);
                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View Scope=\"RecursiveAll\"><Query></Query><RowLimit>5000</RowLimit></View>";
                countImgPdf = 0;
                do
                { 
                    Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(query);
                    context.Load(items, i=>i.Include(itm => itm.FileSystemObjectType), i => i.ListItemCollectionPosition);
                    context.ExecuteQuery();
                   
                    foreach (Microsoft.SharePoint.Client.ListItem listItem in items)
                    {
                        if (listItem.FileSystemObjectType.ToString() != "Folder")
                        {
                            context.Load(listItem, i => i.File);
                            context.ExecuteQuery();
                            if (listItem.File.Name.Substring(listItem.File.Name.LastIndexOf('.') + 1).ToUpper() == "PDF")
                            {
                                //ClientResult<System.IO.Stream> data = listItem.File.OpenBinaryStream();
                                FileInformation data = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, listItem.File.ServerRelativeUrl);
                                context.Load(listItem.File);
                                context.ExecuteQuery();

                                objPDF.fileName = listItem.File.Name;
                                objPDF.fileUrl = listItem.File.ServerRelativeUrl;
                                Console.WriteLine(objPDF.fileName);
                                getImagePdf(data);

                            }
                        }
                    }
                    query.ListItemCollectionPosition = items.ListItemCollectionPosition;

                } while (query.ListItemCollectionPosition != null);


            }
            catch(Exception error)
            {
                Console.WriteLine(error.Message);
            }
        }

        public static void ExportFilesToExcel()
        {
 
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "PDFImageCount";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "PDF with Image";
                worKsheeT.Cells.Font.Size = 15;


                int rowcount = 2;

                foreach (DataRow datarow in table.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= table.Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = table.Columns[i - 1].ColumnName;
                            //worKsheeT.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == table.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, table.Columns.Count]];
                                }

                            }
                        }

                    }

                }

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, table.Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, table.Columns.Count]];

                dir.Create();
                worKbooK.SaveAs(@"C:\PDF Reports\PDFImage_"+ DateTime.Now.Ticks.ToString() + ".xlsx");
                worKbooK.Close();
                excel.Quit();
                Console.WriteLine("File Created at: "+@"C:\PDF Reports\PDFImage_" + DateTime.Now.Ticks.ToString() + ".xlsx");
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);

            }
            finally
            {
                worKsheeT = null;
                celLrangE = null;
                worKbooK = null;
            }
        }



        public static void getImagePdf(FileInformation data)
        {
            string textPDF = string.Empty;
            using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
            {
                if (data != null)
                {
                    data.Stream.CopyTo(mStream);
                    byte[] array = mStream.ToArray();
                    PdfReader reader = new PdfReader(array);
                    int n = reader.XrefSize;
                    PdfObject po;
                    PRStream pst;
                    PdfImageObject pio;
                    FileStream fs = null;
                    //String path2 = "D:/imagesextracted/";

                    for (int i = 0; i < n; i++)
                    {
                        po = reader.GetPdfObject(i); //get the object at the index i in the objects collection
                        if (po == null || !po.IsStream()) //object not found so continue
                            continue;
                        pst = (PRStream)po; //cast object to stream
                        PdfObject type = pst.Get(PdfName.SUBTYPE); //get the object type
                                                                   //check if the object is the image type object
                        if (type != null && type.ToString().Equals(PdfName.IMAGE.ToString()))
                        {
                            //pio = new PdfImageObject(pst); //get the image
                            //fs = new FileStream(path2 + "image" + i + ".jpg", FileMode.Create);
                            ////read bytes of image in to an array
                            //byte[] imgdata = pio.GetImageAsBytes();
                            ////write the bytes array to file
                            //fs.Write(imgdata, 0, imgdata.Length);
                            //fs.Flush();
                            //fs.Close();
                            countImgPdf++;

                            imgPDF.Add(new PDFFile() { fileName=objPDF.fileName,fileUrl=objPDF.fileUrl });
                            table.Rows.Add(objPDF.fileName,objPDF.fileUrl);
                            break;
                        }
                    }
                    reader.Close();
                }
            }
        }



        public static void createPdfFiles()
        {
            Web web = context.Web;
            var list = web.Lists.GetByTitle("PDFFiles");
            context.Load(list.RootFolder);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.ListItem item = list.GetItemById(1);
            context.Load(item);
            context.ExecuteQuery();
            Microsoft.SharePoint.Client.File file = item.File;
            context.Load(item);
            context.Load(file);
            context.ExecuteQuery();
            FileInformation data = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, item.File.ServerRelativeUrl);
            context.Load(item.File);
            context.ExecuteQuery();


            MemoryStream memStream = new MemoryStream();
            data.Stream.CopyTo(memStream);
            var docStream = memStream.ToArray();
            


            for (int i = 0; i <= 5; i++)
            {
                MemoryStream stream = new MemoryStream(docStream);
                var fileUrl = System.IO.Path.Combine(list.RootFolder.ServerRelativeUrl, "Test" + count + ".pdf");
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, stream, true);
                Console.WriteLine("Test" + count + ".pdf");
                count++;
            }
        }
    }
}
