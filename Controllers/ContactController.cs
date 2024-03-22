using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using QRCoder;
using QRVCard.Models;
using System.Drawing;

namespace QRVCard.Controllers
{
    public class ContactController : Controller
    {
        static ContactController()
        {            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult GenerateQrCodes(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                ModelState.AddModelError(string.Empty, "Please select an Excel file.");
                return View("Index");
            }

            List<Contact> contacts = ReadContactsFromExcel(excelFile);

            if (contacts.Count == 0)
            {
                ModelState.AddModelError(string.Empty, "No contacts found in the Excel file.");
                return View("Index");
            }

            List<byte[]> qrCodeImages = GenerateQrCodesForContacts(contacts);

            // Pass the list of QR code images along with contact names to the view
            ViewBag.QrCodeImages = qrCodeImages;
            ViewBag.ContactNames = contacts.Select(c => $"{c.FirstName} {c.LastName}").ToList();

            return View("QrCodeDisplay", qrCodeImages);
        }

        private List<Contact> ReadContactsFromExcel(IFormFile excelFile)
        {
            var contacts = new List<Contact>();

            using (var stream = new MemoryStream())
            {
                excelFile.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming the contacts are in the first sheet
                    if(worksheet.Dimension == null)
                    {
                        return contacts;
                    }
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                    {
                        var contact = new Contact
                        {
                            FirstName = worksheet.Cells[row, 1].Value?.ToString(),
                            LastName = worksheet.Cells[row, 2].Value?.ToString(),
                            Mobile = worksheet.Cells[row, 3].Value?.ToString(),
                            PhoneNumber = worksheet.Cells[row, 4].Value?.ToString(),
                            Email = worksheet.Cells[row, 5].Value?.ToString(),
                            Company = worksheet.Cells[row, 6].Value?.ToString(),
                            Website = worksheet.Cells[row, 7].Value?.ToString()
                        };

                        contacts.Add(contact);
                    }
                }
            }

            return contacts;
        }

        private List<byte[]> GenerateQrCodesForContacts(List<Contact> contacts)
        {
            var qrCodeImages = new List<byte[]>();

            foreach (var contact in contacts)
            {
                var vCardText = $"BEGIN:VCARD\n" +
                                $"VERSION:3.0\n" +
                                $"N:{contact.LastName};{contact.FirstName};;;\n" +
                                $"FN:{contact.FirstName} {contact.LastName}\n" +
                                $"TEL;TYPE=cell:{contact.Mobile}\n" +
                                $"TEL;TYPE=work:{contact.PhoneNumber}\n" +
                                $"EMAIL:{contact.Email}\n" +
                                $"ORG:{contact.Company}\n" +
                                $"URL:{contact.Website}\n" +
                                $"END:VCARD";

                using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(vCardText, QRCodeGenerator.ECCLevel.Q))
                using (QRCode qrCode = new QRCode(qrCodeData))
                using (Bitmap qrCodeImage = qrCode.GetGraphic(10))
                using (var ms = new MemoryStream())
                {
                    qrCodeImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    qrCodeImages.Add(ms.ToArray());
                }
            }

            return qrCodeImages;
        }
    }
}
