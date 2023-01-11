using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Data;
using TestApp.Models;
using System.Diagnostics;
using System.Reflection;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace TestApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {

            List<TestModel> lista = new List<TestModel>();
            TestModel test1 = new TestModel
            {
                Id = 1,
                Name = "Test1",
                Surname = "Test1",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 1,
                TableControl = "asdasd",
                Explanation = 1,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code"

            };

            TestModel test2 = new TestModel
            {
                Id = 2,
                Name = "Test2",
                Surname = "Test2",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 2,
                TableControl = "asdasd",
                Explanation = 2,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code"


            };

            TestModel test3 = new TestModel
            {
                Id = 3,
                Name = "Test3",
                Surname = "Test3",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 3,
                TableControl = "asdasd",
                Explanation = 3,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code"


            };

            TestModel test4 = new TestModel
            {
                Id = 4,
                Name = "Test4",
                Surname = "Te4st4",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 4,
                TableControl = "asdasd",
                Explanation = 4,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code",



            };

            lista.Add(test1);
            lista.Add(test2);
            lista.Add(test3);
            lista.Add(test4);

            return View(lista);
        }

        [HttpPost]
        public IActionResult DownloadReport()
        {
            string reportname = $"User_Wise_{Guid.NewGuid():N}.xlsx";
            List<TestModel> lista= new List<TestModel>();
            TestModel test1 = new TestModel
            {
                Id = 1,
                Name= "Test1",
                Surname= "Test1",
                StartDate= DateTime.Now,
                FinishDate= DateTime.Now,
                DealerList ="asdasdasd",
                FuelList= "asdasdasd",
                Control=1,
                TableControl="asdasd",
                Explanation=1,
                ExplanationTypes="asdasdasd",
                PageType="asdasd",
                DataCompany="asdasdasd",
                DataSelectCompany="asdasd",
                FuelCode="code"
                
            };

            TestModel test2 = new TestModel
            {
                Id = 2,
                Name = "Test2",
                Surname = "Test2",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 2,
                TableControl = "asdasd",
                Explanation = 2,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code"


            };

            TestModel test3 = new TestModel
            {
                Id = 3,
                Name = "Test3",
                Surname = "Test3",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 3,
                TableControl = "asdasd",
                Explanation = 3,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code"


            };

            TestModel test4 = new TestModel
            {
                Id = 4,
                Name = "Test4",
                Surname = "Te4st4",
                StartDate = DateTime.Now,
                FinishDate = DateTime.Now,
                DealerList = "asdasdasd",
                FuelList = "asdasdasd",
                Control = 4,
                TableControl = "asdasd",
                Explanation = 4,
                ExplanationTypes = "asdasdasd",
                PageType = "asdasd",
                DataCompany = "asdasdasd",
                DataSelectCompany = "asdasd",
                FuelCode = "code",
                


            };

            lista.Add(test1);
            lista.Add(test2);
            lista.Add(test3);
            lista.Add(test4);


           var  list=lista.ToList<object>();


            DataTable dat1 = new DataTable("Ue1");

            dat1 = DataTable_Stok_Example();

            var ext = "pdf";
            //var ext = (model.DownloadFileType == 1 || model.DownloadFileType == 2)
            //? "xlsx"
            //: (model.DownloadFileType == 3 || model.DownloadFileType == 4) ? "pdf"
            //: (model.DownloadFileType == 5 || model.DownloadFileType == 6) ? "csv" : "xls";


            byte[] array = PdfExport(new DataSet
            {
                Tables =
                            {
                                dat1
                            }
            }, dat1, " - GENEL KARŞILAŞTIRMALI ÖZET");


            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.Headers.Add("Content-Disposition", "attachment;filename=" + Uri.EscapeUriString($"Ue4T_{DateTime.Now:dd-MMM-yyyy HH-mm}") +
                                    "." + ext);
            Response.Body.Write(array);
            Response.Body.Flush();

            return View();
        }
        public static DataTable DataTable_Stok_Example()
        {
            DataTable table = new DataTable("Stoklar");
            table.Columns.Add(new DataColumn("StokID", typeof(int)));
            table.Columns.Add(new DataColumn("StokKodu", typeof(string)));
            table.Columns.Add(new DataColumn("StokAdi", typeof(string)));
            table.Columns.Add(new DataColumn("StokBirimi", typeof(string)));
            table.Columns.Add(new DataColumn("StokKDVOran", typeof(int)));
            table.Columns.Add(new DataColumn("StokBirimFiyat", typeof(double)));
            table.Columns.Add(new DataColumn("StokGrupKodu", typeof(string)));

            table.Rows.Add(1, "şüğçİıĞÜŞ", "Klavye", "Adet", 18, 65, "PC");
            table.Rows.Add(2, "şüğçİı", "Kablolu Mause", "Adet", 18, 50, "PC");
            table.Rows.Add(3, "şüğçİı", "20 inc Monitör", "Adet", 18, 225, "PC");
            table.Rows.Add(4, "S004", "Kaşüğçsa", "Adet", 18, 80, "PC");
            table.Rows.Add(5, "S005", "Kaşüğç400 Watt PowerSupply", "Adet", 18, 120, "PC");
            table.Rows.Add(6, "S006", "StereKaşüğço Hoparlor", "Adet", 18, 55, "PC");
            table.Rows.Add(7, "S007", "KulaklıKaşüğçİık", "Adet", 18, 60, "PC");
            table.Rows.Add(8, "S008", "CAT 6 Kablo", "Metre", 18, 1.75, "KAblo");
            table.Rows.Add(9, "S009", "Kablosuz Mause", "Adet", 18, 65, "PC");
            table.Rows.Add(10, "S010", "CAT 5 Kablo", "Metre", 18, 1.25, "Kablo");
            table.Rows.Add(11, "S011", "Kablosuz Klavye", "Adet", 18, 85, "PC");
            return table;
        }
        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
        public static byte[] PdfExport(DataSet dsInput, DataTable orjDtInput, string caption)
        {

            Document val = new Document(PageSize.A4_LANDSCAPE, 5f, 5f, 5f, 5f);
            MemoryStream memoryStream = new MemoryStream();
            byte[] result = new byte[0];
            try
            {
                PdfWriter.GetInstance(val, (Stream)memoryStream);
                val.Open();
                Font currentFont = new Font( BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false));

                Font currentFontRed = new Font(BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false));
                currentFontRed.Color = BaseColor.RED;
                if (dsInput != null && dsInput.Tables.Count > 0)
                {
                    PdfPTable title = new PdfPTable(1);
                    var titleCell = new PdfPCell(new Phrase(new Chunk(caption, new Font(BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false)))));
                    title.AddCell(titleCell);
                    val.Add(title);
                    foreach (DataTable table in dsInput.Tables)
                    {
                        PdfPTable val3 = new PdfPTable(table.Columns.Count);
                        PdfPCell val4 = null;
                        foreach (DataColumn column in table.Columns)
                        {

                            val4 = new PdfPCell(new Phrase(new Chunk(string.IsNullOrWhiteSpace(column.Caption) ? column.ColumnName.ToString() : column.Caption.ToString(), currentFont)));
                            val3.AddCell(val4);
                        }
                        int num = -1;
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                var columnCurrent = table.Columns[j];
                                var r = table.Rows[i];
                                var o = r[j];
                                var valTemp = (o == null ? "" : o);
                                val4 = new PdfPCell(new Phrase(new Chunk(valTemp.ToString(), currentFont)));
                                if (table.TableName == "Ue4T")
                                {
                                    #region Stok Açılış veya Kapanış  
                                    if (j == 0)
                                    {
                                        if (valTemp.ToString().Contains("00:00:00") && valTemp.ToString().ToString()
                                                   .Split(' ').Length > 0)
                                        {
                                            val4 = new PdfPCell(new Phrase(new Chunk(valTemp.ToString().Split(' ')[0], currentFont)));
                                        }
                                    }
                                    if (j == 3) //Tankların Manifold Değeri
                                    {
                                        string connPumps = table.Rows[i][5].ToString();
                                        if (connPumps.Contains("*"))
                                        {
                                            val4.BackgroundColor = BaseColor.RED;
                                        }

                                    }
                                    else if (j == 5) //Pompa Tabanca Ayarı
                                    {
                                        if (valTemp.ToString().Contains("*"))
                                        {
                                            val4 = new PdfPCell(new Phrase(new Chunk(valTemp.ToString().Replace("*", ""), currentFont)));
                                        }
                                        if (valTemp.ToString() == "" || valTemp == null)
                                        {

                                            val4 = new PdfPCell(new Phrase(new Chunk("Yedek Tanka Bağlı Tabanca Yok", currentFontRed)));

                                        }

                                    }
                                    else if (j == 6 || j == 7)//"Stok Açılış""Stok Kapanış"
                                    {

                                        // Not:TODO:::Elpo Değilse İşlem Yapma ,Linecomment alanına göre Elpo olan verilerde güncelleme yapılacak Offline,Tadilat,Arıza bilgisini yazdırabiliyor olucak
                                        string LineInfo = table.Rows[i][14].ToString();
                                        string firm = table.Rows[i][15].ToString();
                                        if (firm != "Elpo")
                                        {
                                            if (valTemp == null || valTemp.ToString() == "")
                                            {
                                                #region MyRegion
                                                var oo = "";
                                                if (LineInfo != null && LineInfo != "" && LineInfo.IndexOf('#') > 0)
                                                {
                                                    var lineOpen = LineInfo.Split('|');
                                                    foreach (var item in lineOpen)
                                                    {
                                                        if (item.IndexOf((j == 6 ? "Opening" : (j == 7) ? "Closing" : "")) >= 0)
                                                        {
                                                            oo = item.Replace((j == 6 ? "Opening#" : (j == 7) ? "Closing#" : ""), "");
                                                            break;
                                                        }
                                                    }
                                                    val4 = new PdfPCell(new Phrase(new Chunk(oo, currentFont)));
                                                }
                                                else
                                                {
                                                    val4 = new PdfPCell(new Phrase(new Chunk("Hatalı", currentFont)));
                                                }
                                                #endregion
                                                val4.BackgroundColor = BaseColor.RED;
                                            }
                                        }
                                        else
                                        {
                                            if (valTemp == null || valTemp.ToString() == "")
                                            {
                                                val4 = new PdfPCell(new Phrase(new Chunk("Hatalı", currentFont)));
                                                val4.BackgroundColor = BaseColor.RED;

                                            }
                                        }
                                    }
                                    else if (j == 10) //Ue-1 e göre satış
                                    {
                                        string firm = table.Rows[i][15].ToString();

                                        if (firm == "Elpo")
                                        {
                                            var opening = table.Rows[i][6];
                                            var closing = table.Rows[i][7];
                                            var filling = table.Rows[i][9];

                                            if (valTemp == null || valTemp.ToString() == "")
                                            {
                                                if (closing != null && opening != null
                                                    && closing.ToString() != "" && opening.ToString() != "")
                                                {
                                                    var ue1Sales = decimal.Parse(opening.ToString()) - decimal.Parse(closing.ToString()) + decimal.Parse(filling.ToString());
                                                    val4 = new PdfPCell(new Phrase(new Chunk(ue1Sales.ToString(), currentFont)));
                                                }
                                                else
                                                {
                                                    val4 = new PdfPCell(new Phrase(new Chunk("Hatalı", currentFont)));
                                                    val4.BackgroundColor = BaseColor.RED;
                                                }
                                            }

                                            decimal tempop;
                                            decimal tempclo;
                                            if (!decimal.TryParse(opening.ToString(), out tempop) && !decimal.TryParse(closing.ToString(), out tempclo))
                                            {
                                                val4 = new PdfPCell(new Phrase(new Chunk("Hatalı", currentFont)));
                                                val4.BackgroundColor = BaseColor.RED;
                                            }
                                        }
                                        else
                                        {
                                            #region MyRegion
                                            string LineInfo = table.Rows[i][14].ToString();
                                            var oo = "";
                                            if (LineInfo != null && LineInfo != "" && LineInfo.IndexOf('#') > 0)
                                            {
                                                var lineOpen = LineInfo.Split('|');
                                                foreach (var item in lineOpen)
                                                {
                                                    if (item.IndexOf("Ue1Sales") >= 0)
                                                    {
                                                        oo = item.Replace("Ue1Sales#", "");
                                                        break;
                                                    }
                                                }
                                                val4 = new PdfPCell(new Phrase(new Chunk(oo, currentFont)));
                                                val4.BackgroundColor = BaseColor.RED;
                                            }
                                            #endregion
                                        }

                                    }
                                    else if (j == 12)
                                    {
                                        string firm = table.Rows[i][15].ToString();

                                        if (firm == "Elpo")
                                        {

                                        }
                                        else
                                        {

                                            #region MyRegion
                                            string LineInfo = table.Rows[i][14].ToString();
                                            var oo = "";
                                            if (LineInfo != null && LineInfo != "" && LineInfo.IndexOf('#') > 0)
                                            {
                                                var lineOpen = LineInfo.Split('|');
                                                foreach (var item in lineOpen)
                                                {
                                                    if (item.IndexOf("ReductionAmount") >= 0)
                                                    {
                                                        oo = item.Replace("ReductionAmount#", "");
                                                        break;
                                                    }
                                                }
                                                if (!String.IsNullOrEmpty(oo))
                                                {
                                                    val4 = new PdfPCell(new Phrase(new Chunk(oo, currentFont)));
                                                    val4.BackgroundColor = BaseColor.RED;
                                                }
                                            }
                                            #endregion
                                        }

                                    }
                                    else if (j == 13)
                                    {
                                        if (valTemp == null || valTemp.ToString() == "")
                                        {
                                            val4 = new PdfPCell(new Phrase(new Chunk("Hatalı", currentFont)));
                                            val4.BackgroundColor = BaseColor.RED;
                                        }
                                    }
                                    else if (j == 14)
                                    {
                                        if (val4 != null || val4.ToString() != "")
                                        {
                                            if (val4.ToString().IndexOf('#') > 0)
                                            {
                                                string oo = val4.ToString().Replace("|Opening#", ",").Replace("|Closing#", ",").Replace("|Ue1Sales#", ",").Replace("|ReductionAmount#", ",");
                                                val4 = new PdfPCell(new Phrase(new Chunk(oo, currentFont)));
                                                val4.BackgroundColor = BaseColor.RED;
                                            }
                                        }
                                    }
                                    #endregion

                                }
                                val3.AddCell(val4);
                            }
                        }
                        val.Add(new Phrase(new Chunk(table.TableName + " Verileri", currentFont)));
                        val.Add(val3);
                        val.NewPage();
                    }
                }
                val.Close();
                result = memoryStream.GetBuffer();
                return result;
            }
            catch (DocumentException)
            {
                return result;
            }
            catch (IOException)
            {
                return result;
            }
            catch (Exception)
            {
                return result;
            }
        }

        private byte[] ExporttoExcel<T>(List<T> table, string filename)
        {
            using ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet ws = pack.Workbook.Worksheets.Add(filename);
            ws.Cells["A1"].LoadFromCollection(table, true, TableStyles.Light1);

            ws = pack.Workbook.Worksheets.Add("asdasdasd");
            ws.Cells["A1"].LoadFromCollection(table, true, TableStyles.Light1);

            ws = pack.Workbook.Worksheets.Add("asdasdww");
            ws.Cells["A1"].LoadFromCollection(table, true, TableStyles.Light1);
            return pack.GetAsByteArray();
        }


        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}