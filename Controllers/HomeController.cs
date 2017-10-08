using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using WebApplication.Models;

namespace WebApplication.Controllers
{
    public class HomeController : Controller
    {


        public ActionResult Index()
        {
            ExampleEntities db = new ExampleEntities();
            
            return View(db.Product);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }



        [HttpPost]
        public ActionResult ExcelUpload(HttpPostedFileBase postedFile)
        {

            string filePath = string.Empty;


            if (postedFile != null)
            {
                string path = Server.MapPath("~/Uploads/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }



                //identifying file extension to decide which Excel version to use as connection string. The string info is in the Web.Config file
                filePath = path + Path.GetFileName(postedFile.FileName);
                string extension = Path.GetExtension(postedFile.FileName);
                postedFile.SaveAs(filePath);

                string conString = string.Empty;
                switch (extension)
                {
                    case ".xls": //Connection string for Excel 97-03.
                        conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                        break;
                    case ".xlsx": //Connection string for Excel 07 and above.
                        conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                        break;
                }

                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {

                            using (ExampleEntities db = new ExampleEntities())
                            {

                                cmdExcel.Connection = connExcel;

                                //Get the name of First Sheet.
                                connExcel.Open();
                                DataTable dtExcelSchema;
                                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                connExcel.Close();

                                //Read Data from First Sheet.
                                connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "] ";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);


                                // This part required for deleting empty columns. Excel file automatically inserts empty rows, if havent been removed, they will be added as null data.
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    if (dt.Rows[i]["Name"].ToString() == null || dt.Rows[i]["Name"].ToString() == "")
                                    {
                                        dt.Rows[i].Delete();
                                    }
                                }
                                dt.AcceptChanges();
                                int newRowCount = dt.Rows.Count;

                                for (int i = 0; i < newRowCount; i++)
                                {
                                    var name = dt.Rows[i]["Name"].ToString();
                                    var info = dt.Rows[i]["Info"].ToString();
                                    var price = dt.Rows[i]["Price"];
                                    if (price.ToString() == "")
                                    {
                                        price = 0;
                                    }
                                    else
                                    {
                                        price = Convert.ToDouble(price);
                                    }

                                    var status = dt.Rows[i]["Status"];
                                    if (status.ToString() == "")
                                    {
                                        status = 0;
                                    }
                                    else
                                    {
                                        status = Convert.ToInt32(status);
                                    }

                                    var registerDate = DateTime.Now;

                                    var stock = dt.Rows[i]["Stock"];
                                    if (stock.ToString() != "")
                                    {
                                        stock = Convert.ToInt32(dt.Rows[i]["Stock"]);
                                    }


                                    Product _product = new Product()
                                    {
                                        Name = name,
                                        Info = info,
                                        RegisterDate = registerDate,
                                        Price = Convert.ToDouble(price),
                                        Status = Convert.ToInt32(status),
                                        Stock = Convert.ToInt32(stock)

                                    };
                                    db.Product.Add(_product);
                                    db.SaveChanges();

                                }
                            }
                            connExcel.Close();
                        }
                    }
                }
            }

            return RedirectToAction("Index");
        }
    }
}