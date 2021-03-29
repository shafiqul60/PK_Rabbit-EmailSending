using EmailSending.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web.Mvc;
using ClosedXML.Excel;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using System.Web;

namespace EmailSending.Controllers
{
    public class MailController : Controller
    {

        ProjectDBContext _db = new ProjectDBContext();

        

        public ActionResult Index()
        {
            return View();
        }
       
        [HttpPost]
        public ActionResult Index(EmailSending.Models.MailModel objModelMail, HttpPostedFileBase fileUploader)
        {
            if (ModelState.IsValid)
            {
                string from = "shafiqul.cse2017@gmail.com"; 
                using (MailMessage mail = new MailMessage(from, objModelMail.To))
                {
                    mail.Subject = objModelMail.Subject;
                    mail.Body = objModelMail.Body;
                    if (fileUploader != null)
                    {
                        string fileName = Path.GetFileName(fileUploader.FileName);
                        mail.Attachments.Add(new Attachment(fileUploader.InputStream, fileName));
                    }
                    mail.IsBodyHtml = false;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;
                    NetworkCredential networkCredential = new NetworkCredential(from, "shafiqul&60");
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = networkCredential;
                    smtp.Port = 587;
                    smtp.Send(mail);
                    ViewBag.Suc = "Message Successfully Sent";

                    MailModel email = new MailModel();

                    email.To = objModelMail.To;
                    
                    email.Subject = objModelMail.Subject;
                    email.Body = objModelMail.Body;
                    _db.MailModels.Add(email);
                    _db.SaveChanges();

                    return View("Index", objModelMail);
                }
            }
            else
            {
                ViewBag.Error = "Message Successfully Sending Failed";
                return View();
            }
        }


        public IList<MailModel> GetMailList()
        {
            var mailList = _db.MailModels.ToList();

            return mailList;
        }

        public ActionResult Details()
        {
            return View(this.GetMailList());
        }


        [HttpPost]
        public FileResult ExportToExcel()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[3] {
                                                     new DataColumn("MailTo"),                                                   
                                                     new DataColumn("MailSubject"),
                                                     new DataColumn("MailBody")});

            var MailList = from email in _db.MailModels select email;

            foreach (var m in MailList)
            {
                dt.Rows.Add(m.To, m.Subject, m.Body);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (MemoryStream stream = new MemoryStream()) 
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExcelFile.xlsx");
                }
            }
        }

        [HttpPost]
        [ValidateInput(false)]
        public FileResult Export(string GridHtml)
        {
            using (MemoryStream stream = new System.IO.MemoryStream())
            {
                StringReader sr = new StringReader(GridHtml);
                Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 100f, 0f);
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, stream);
                pdfDoc.Open();
                XMLWorkerHelper.GetInstance().ParseXHtml(writer, pdfDoc, sr);
                pdfDoc.Close();
                return File(stream.ToArray(), "application/pdf", "PdfFile.pdf");
            }
        }




    }
}