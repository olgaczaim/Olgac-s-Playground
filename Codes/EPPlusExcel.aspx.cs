using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication13
{
    public partial class WebForm4 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //ExcelPackage.LicenseContext = LicenseContext.Commercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "VDWWD";
                excelPackage.Workbook.Properties.Title = "Title of Document";
                excelPackage.Workbook.Properties.Subject = "EPPlus demo export data";
                excelPackage.Workbook.Properties.Created = DateTime.Now;
                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Tetkik Raporu");
                //Add some text to cell A1
                worksheet.Cells["A1:B2"].Merge = true;
                worksheet.Cells["A1:B2"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A1:B2"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A1:B2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A1:B2"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                
                var cell = worksheet.Cells["C1:K2"];
                cell.Merge = true;
                var cellFont = cell.Style.Font;
                cellFont.SetFromFont(new Font("Arial", 18));
                var title = cell.RichText.Add("...Tetkik Raporu");
                title.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell3 = worksheet.Cells["L1"];
                var cellFont2 = cell3.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell3.RichText.Add("Tetkik Ref. no");
                title.Bold = true;
                //cell3.AutoFitColumns();
                cell3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell3.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell3.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell3.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell3.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell4 = worksheet.Cells["L2"];
                cell4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell4.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell4.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell4.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell4.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell5 = worksheet.Cells["A3:B3"];
                cell5.Merge = true;
                cellFont2 = cell5.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell5.RichText.Add("Tetkikin Amacı");
                title.Bold = true;
                cell5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell5.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell5.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell5.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell5.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell5.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell6 = worksheet.Cells["C3:J3"];
                cell6.Merge = true;
                cellFont2 = cell6.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell6.RichText.Add("sec1");
                cell6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell6.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell6.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell6.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell6.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell6.Style.Border.Left.Style = ExcelBorderStyle.Thin;


                var cell7 = worksheet.Cells["K3"];
                cellFont2 = cell7.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell7.RichText.Add("Tetkik Tarihi");
                title.Bold = true;
                //cell7.AutoFitColumns();
                cell7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell7.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell7.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell7.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell7.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell7.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell8 = worksheet.Cells["L3"];
                cellFont2 = cell8.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell8.RichText.Add("sec2");
                //cell8.AutoFitColumns();
                cell8.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell8.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell8.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell8.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell8.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell8.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell9 = worksheet.Cells["A4:B4"];
                cell9.Merge = true;
                cellFont2 = cell9.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell9.RichText.Add("Tetkik Kriterleri");
                title.Bold = true;
                cell9.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell9.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell9.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell9.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell9.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell9.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell10 = worksheet.Cells["C4:G4"];
                cell10.Merge = true;
                cellFont2 = cell10.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell10.RichText.Add("sec3");
                //cell10.AutoFitColumns();
                cell10.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell10.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell10.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell10.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell10.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell10.Style.Border.Left.Style = ExcelBorderStyle.Thin;


                var cell11 = worksheet.Cells["H4:J4"];
                cell11.Merge = true;
                cellFont2 = cell11.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell11.RichText.Add("Tetkik Edilen Birim");
                title.Bold = true;
                cell11.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell11.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell11.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell11.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell11.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell11.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell12 = worksheet.Cells["K4:L4"];
                cell12.Merge = true;
                cellFont2 = cell12.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell12.RichText.Add("sec4");
                //cell12.AutoFitColumns();
                cell12.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell12.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell12.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell12.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell12.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell12.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                List<User> listOfUsers = new List<User>()
                {
                    new User() { adSoyad = "John Doe", unvan = "Denetçi", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Jane Doe", unvan = "Yazılım Uzmanı", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Jane Doe", unvan = "Yazılım Uzmanı", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Jane Doe", unvan = "Yazılım Uzmanı", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Jane Doe", unvan = "Yazılım Uzmanı", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Joe Doe", unvan = "Bilgi İşlem Müdürü", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                };

                int tetkikciCount = listOfUsers.Count() + 1 + 4;

                var cell13 = worksheet.Cells["A5:B" + tetkikciCount + ""];
                cell13.Merge = true;
                cellFont2 = cell13.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell13.RichText.Add("Tetkikçi(ler)");
                title.Bold = true;
                cell13.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell13.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell13.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell13.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell13.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell13.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                var cell14 = worksheet.Cells["C5:E5"];
                cell14.Merge = true;
                cellFont2 = cell14.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell14.RichText.Add("AD SOYAD");
                title.Bold = true;
                cell14.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell14.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell14.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell14.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell14.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell14.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["C5:E5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["C5:E5"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                var cell15 = worksheet.Cells["F5:H5"];
                cell15.Merge = true;
                cellFont2 = cell15.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell15.RichText.Add("UNVAN");
                title.Bold = true;
                cell15.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell15.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell15.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell15.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell15.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell15.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["F5:H5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["F5:H5"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                var cell16 = worksheet.Cells["I5:J5"];
                cell16.Merge = true;
                cellFont2 = cell16.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell16.RichText.Add("TARIH");
                title.Bold = true;
                cell16.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell16.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell16.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell16.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell16.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell16.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["I5:J5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["I5:J5"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                var cell17 = worksheet.Cells["K5:L5"];                
                cell17.Merge = true;
                cellFont2 = cell17.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell17.RichText.Add("IMZA");
                title.Bold = true;
                cell17.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell17.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell17.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell17.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell17.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell17.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["K5:L5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["K5:L5"].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                

                int curr = 5;
                foreach (User item in listOfUsers)
                {
                    curr++;
                    var cellTekAdi = worksheet.Cells["C" + curr + ":E" + curr + ""];
                    cellTekAdi.Merge = true;
                    cellFont2 = cellTekAdi.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellTekAdi.RichText.Add(item.adSoyad);
                    cellTekAdi.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellTekAdi.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellTekAdi.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellTekAdi.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellTekAdi.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellTekAdi.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    var cellUnvan = worksheet.Cells["F" + curr + ":H" + curr + ""];
                    cellUnvan.Merge = true;
                    cellFont2 = cellUnvan.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellUnvan.RichText.Add(item.unvan);
                    cellUnvan.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellUnvan.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellUnvan.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellUnvan.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellUnvan.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellUnvan.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    var cellTarih = worksheet.Cells["I" + curr + ":J" + curr + ""];
                    cellTarih.Merge = true;
                    cellFont2 = cellTarih.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellTarih.RichText.Add(item.tarih);
                    cellTarih.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellTarih.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellTarih.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellTarih.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellTarih.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellTarih.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    var cellImza = worksheet.Cells["K" + curr + ":L" + curr + ""];
                    cellImza.Merge = true;
                    cellFont2 = cellImza.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellImza.RichText.Add(" ");
                    cellImza.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellImza.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellImza.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellImza.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellImza.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellImza.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                }

                List<User> listGorusulenler = new List<User>()
                {
                    new User() { adSoyad = "John Doe", unvan = "Denetçi", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Jane Doe", unvan = "Yazılım Uzmanı", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Jane Doe", unvan = "Yazılım Uzmanı", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                    new User() { adSoyad = "Joe Doe", unvan = "Bilgi İşlem Müdürü", tarih = DateTime.Now.ToString("dd.MM.yyyy"), imza = "" },
                };

                var cellGorusulenler = worksheet.Cells["A" + (curr + 1) + ":B" + (curr + listGorusulenler.Count()) + ""];
                cellGorusulenler.Merge = true;
                cellFont2 = cellGorusulenler.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cellGorusulenler.RichText.Add("Görüşülen(ler)");
                title.Bold = true;
                cellGorusulenler.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cellGorusulenler.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cellGorusulenler.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cellGorusulenler.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cellGorusulenler.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cellGorusulenler.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                foreach (var item in listGorusulenler)
                {
                    curr++;
                    var cellTekAdi = worksheet.Cells["C" + curr + ":E" + curr + ""];
                    cellTekAdi.Merge = true;
                    cellFont2 = cellTekAdi.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellTekAdi.RichText.Add(item.adSoyad);
                    cellTekAdi.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellTekAdi.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellTekAdi.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellTekAdi.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellTekAdi.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellTekAdi.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    var cellUnvan = worksheet.Cells["F" + curr + ":H" + curr + ""];
                    cellUnvan.Merge = true;
                    cellFont2 = cellUnvan.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellUnvan.RichText.Add(item.unvan);
                    cellUnvan.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellUnvan.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellUnvan.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellUnvan.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellUnvan.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellUnvan.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    var cellTarih = worksheet.Cells["I" + curr + ":J" + curr + ""];
                    cellTarih.Merge = true;
                    cellFont2 = cellTarih.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellTarih.RichText.Add(item.tarih);
                    cellTarih.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellTarih.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellTarih.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellTarih.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellTarih.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellTarih.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    var cellImza = worksheet.Cells["K" + curr + ":L" + curr + ""];
                    cellImza.Merge = true;
                    cellFont2 = cellImza.Style.Font;
                    cellFont2.SetFromFont(new Font("Arial", 9));
                    title = cellImza.RichText.Add(" ");
                    cellImza.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cellImza.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellImza.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cellImza.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cellImza.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cellImza.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                }
                curr++;
                var cell18 = worksheet.Cells["A" + (curr) + ":L" + (curr) + ""];
                cell18.Merge = true;
                cellFont2 = cell18.Style.Font;
                cellFont2.SetFromFont(new Font("Arial", 9));
                title = cell18.RichText.Add("Tetkik bulguları 4 ana gurupta sınıflandırılır.");
                title.Italic = true;
                title = cell18.RichText.Add(" 1)Uygunsuzluj, 2)Gözlem 3)Tavsiye 4)Pozitif Bulgu");
                title.Bold = true;
                title.Italic = true;
                cell18.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell18.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell18.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell18.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell18.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell18.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A" + (curr) + ":L" + (curr) + ""].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A" + (curr) + ":L" + (curr) + ""].Style.Fill.BackgroundColor.SetColor(Color.Silver);
                
                curr++;

                worksheet.Column(11).Width = 11;
                worksheet.Column(12).Width = 12;
                //worksheet.Cells["C5"].AutoFilter = true;
                //worksheet.Cells["F5"].AutoFilter = true;
                //worksheet.Cells["I5"].AutoFilter = true;
                //worksheet.Cells["K5"].AutoFilter = true;


                // worksheet.Cells["K4:L4"].Style.WrapText = true;
                //worksheet.Cells["L1"].AutoFitColumns();
                //worksheet.Cells["L3"].AutoFitColumns();
                //worksheet.Cells["K3"].AutoFitColumns();
                worksheet.Row(3).Height = 30;
                worksheet.Row(4).Height = 30;
                //byte[] bin = excelPackage.GetAsByteArray();
                //Response.ClearHeaders();
                //Response.Clear();
                //Response.Buffer = true;
                //Response.ContentType = "application/vnd.openxmlformatsofficedocument.spreadsheetml.sheet";
                ////set the correct length of the data being send
                //Response.AddHeader("content-length", bin.Length.ToString());
                ////set the filename for the excel package
                //Response.AddHeader("content-disposition", "attachment; filename=\"ExcelDemo.xlsx\"");
                ////send the byte array to the browser
                //Response.OutputStream.Write(bin, 0, bin.Length);
                ////cleanup
                //Response.Flush();
                //HttpContext.Current.ApplicationInstance.CompleteRequest();


                FileInfo fi = new FileInfo(@"C:\log\File.xlsx");
                excelPackage.SaveAs(fi);
            }



        }

    }

    class User
    {
        public string adSoyad { get; set; }
        public string unvan { get; set; }
        public string tarih { get; set; }
        public string imza { get; set; }
    }
}
