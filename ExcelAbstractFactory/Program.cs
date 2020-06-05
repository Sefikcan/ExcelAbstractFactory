using ExcelAbstractFactory.Data;
using ExcelAbstractFactory.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace ExcelAbstractFactory
{
    class Program
    {
        static void Main(string[] args)
        {
            List<EmailAttachModel> emailAttachModels = new List<EmailAttachModel>();

            #region tek sheet için

            //var dataList = LocalData.GetProductList();

            //var excelByteModel = new ExcelByteModel<Product>()
            //{
            //    ExcelItemList = dataList,
            //    SheetName = "Sheet1"
            //};

            //IExcelFactory factory = ExcelFactory(ExcelBrandType.OpenXml);

            //var excelbytes = factory.CreateExcelBytes();
            //var byteData = excelbytes.ToExcelBytesOneSheet(excelByteModel);
            //SendEmail(byteData);
            #endregion

            #region çoklu sheet için
            var dataList1 = LocalData.GetProductList();
            var dataList2 = LocalData.GetProductList2();

            List<ExcelByteModel<Product>> excelByteModels = new List<ExcelByteModel<Product>>();
            var excelByte1 = new ExcelByteModel<Product>()
            {
                ExcelItemList = dataList1,
                SheetName = "Sheet1"
            };
            excelByteModels.Add(excelByte1);
            var excelByte2 = new ExcelByteModel<Product>()
            {
                ExcelItemList = dataList2,
                SheetName = "Sheet2"
            };
            excelByteModels.Add(excelByte2);

            IExcelFactory factory = ExcelFactory(ExcelBrandType.OpenXml);
            var excelbytes = factory.CreateExcelBytes();
            var byteData = excelbytes.ToExcelBytesMultiSheets(excelByteModels);

            EmailAttachModel emailAttachModel = new EmailAttachModel
            {
                ByteData = byteData,
                FullFileName = "ExcelSheet1.xlsx", //queueu ile buralar generic yapıya alınabilir
                MediaType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            };

            emailAttachModels.Add(emailAttachModel);

            #endregion


            SendEmail(emailAttachModels);
        }

        static void SendEmail(List<EmailAttachModel> emailAttachModels)
        {
            var msList = new List<MemoryStream>();

            using (SmtpClient client = new SmtpClient())
            using (MailMessage mail = new MailMessage())
            {
                client.Host = "smtp.gmail.com";
                client.Port = 587;
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential("xxx@gmail.com", "xxxx");

                mail.From = new MailAddress("sxt@gmail.com", "test");
                mail.To.Add(new MailAddress("xxx@outlook.com"));
                mail.Subject = "Send excel email attachment c#";
                mail.IsBodyHtml = true;
                mail.Body = "<html><head></head><body>Attached is the Excel sheet.</body></html>";

                //attach the excel file to the message
                //mail.Attachments.Add(new Attachment(ms, "ExcelSheet1.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

                foreach (var emailAttachModel in emailAttachModels)
                {
                    var ms = new MemoryStream(emailAttachModel.ByteData);
                    msList.Add(ms);

                    var mailAttachment = new Attachment(ms, emailAttachModel.FullFileName, emailAttachModel.MediaType);
                    mail.Attachments.Add(mailAttachment);
                }


                //send the mail
                try
                {
                    client.Send(mail);
                }
                catch (Exception ex)
                {
                    //handle error
                }
            }
            foreach (var ms in msList)
            {
                ms.Dispose();
            }
        }

        static IExcelFactory ExcelFactory(ExcelBrandType excelBrandType)
        {
            IExcelFactory factory;
            switch (excelBrandType)
            {
                case ExcelBrandType.OpenXml:
                    factory = new OpenXmlFactory();
                    break;
                case ExcelBrandType.NPOI:
                    factory = new NPOIFactory();
                    break;
                default:
                    throw new System.NotImplementedException();
            }
            return factory;
        }
    }
}
