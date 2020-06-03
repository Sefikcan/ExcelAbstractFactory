using DocumentFormat.OpenXml.Office2010.ExcelAc;
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
            Console.WriteLine("Hello World!");

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
            SendEmail(byteData);
            #endregion
        }

        static void SendEmail(byte[] byteData)
        {
            MemoryStream ms = new MemoryStream(byteData);

            using (SmtpClient client = new SmtpClient())
            using (MailMessage mail = new MailMessage())
            {
                client.Host = "smtp.gmail.com";
                client.Port = 587;
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential("sefikcankanbertest@gmail.com", "q1w2e3r4T5");

                mail.From = new MailAddress("sefikcankanbertest@gmail.com", "test");
                mail.To.Add(new MailAddress("sefikcankanber@outlook.com"));
                mail.Subject = "Send excel email attachment c#";
                mail.IsBodyHtml = true;
                mail.Body = "<html><head></head><body>Attached is the Excel sheet.</body></html>";

                //attach the excel file to the message
                mail.Attachments.Add(new Attachment(ms, "ExcelSheet1.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

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
            ms.Dispose();
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
