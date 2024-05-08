
 
using iTextSharp.text;

using iTextSharp.text.pdf;

using iTextSharp.text.pdf.security;

using Org.BouncyCastle.X509;

using System;

using System.Collections.Generic;

using System.Diagnostics;

using System.IO;

using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;

namespace MPODigiSign

{

    internal class SignHelper

    {

        private int num = 0;

        private string sourcePath = "";

        private string destinationPath = "";

        private string TokenCertificateName = "";

        private string reason = "";

        private string location = "";

        private bool SignOnEveryPage = false;

        private string OutputFileName = "";

        public void SignDocument()

        {

            Console.WriteLine("Reading File Path");

            this.SetFilePath();

            this.SignPdf();

        }

        public string getBetween(string strSource, string strStart, string strEnd)

        {

            if (!strSource.Contains(strStart) || !strSource.Contains(strEnd))

                return "";

            int startIndex = strSource.IndexOf(strStart, 0) + strStart.Length;

            int num = strSource.IndexOf(strEnd, startIndex);

            return strSource.Substring(startIndex, num - startIndex);

        }

        private void SignPdf()

        {

            Console.WriteLine("Searching for certificate");

            X509Store x509Store = new X509Store(StoreName.My, StoreLocation.CurrentUser);

            x509Store.Open(OpenFlags.OpenExistingOnly);

            X509Certificate2Collection certificates = x509Store.Certificates;

            X509Certificate2 x509Certificate2 = new X509Certificate2();

            for (int index1 = 0; index1 < certificates.Count; ++index1)

            {

                X509Certificate2 cert = certificates[index1];

                string subject = cert.Subject;

                List<X509KeyUsageExtension> list = cert.Extensions.OfType<X509KeyUsageExtension>().ToList<X509KeyUsageExtension>();

                if (cert.GetNameInfo(X509NameType.SimpleName, false).ToString() == this.TokenCertificateName)

                {

                    for (int index2 = 0; index2 < list.Count; ++index2)

                    {

                        if ((list[index2].KeyUsages & X509KeyUsageFlags.DigitalSignature) == X509KeyUsageFlags.DigitalSignature)

                        {

                            Console.WriteLine("Signing file...");

                            foreach (string file in Directory.GetFiles(this.sourcePath))

                            {

                                this.OutputFileName = Path.GetFileName(file);

                                if (!this.SignOnEveryPage)

                                {

                                    this.SignWithThisCert(cert, this.location, file, this.destinationPath + this.OutputFileName);

                                    new Process()

                                    {

                                        StartInfo = {

                      FileName = (this.destinationPath + this.OutputFileName)

                    }

                                    }.Start();

                                }

                                else

                                {

                                    this.splitdocument(file);

                                    string path = "C:\\DigitalSignatureSettings\\Split\\";

                                    string[] files = Directory.GetFiles(path);

                                    for (int index3 = 0; index3 < files.Length; ++index3)

                                    {

                                        string SourcePdfFileName = files[index3];

                                        this.destinationPath = path + Path.GetFileNameWithoutExtension(files[index3]) + "Signed.pdf";

                                        this.SignWithThisCert(cert, this.location, SourcePdfFileName, this.destinationPath);

                                        File.Delete(files[index3]);

                                    }

                                    this.mergedocument();

                                }

                            }

                        }

                    }

                }

            }

        }

        private void SignWithThisCert(

          X509Certificate2 cert,

          string locationName,

          string SourcePdfFileName,

          string DestPdfFileName)

        {

            Org.BouncyCastle.X509.X509Certificate[] chain = new Org.BouncyCastle.X509.X509Certificate[1]

            {

        new X509CertificateParser().ReadCertificate(cert.RawData)

            };

            IExternalSignature externalSignature = (IExternalSignature)new X509Certificate2Signature(cert, "SHA-1");

            PdfSignatureAppearance signatureAppearance = PdfStamper.CreateSignature(new PdfReader(SourcePdfFileName), (Stream)new FileStream(DestPdfFileName, FileMode.Create), char.MinValue).SignatureAppearance;

            signatureAppearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.NAME_AND_DESCRIPTION;

            signatureAppearance.SetVisibleSignature(new Rectangle(510f, 190f, 280f, 240f), 1, (string)null);


            MakeSignature.SignDetached(signatureAppearance, externalSignature, (ICollection<Org.BouncyCastle.X509.X509Certificate>)chain, (ICollection<ICrlClient>)null, (IOcspClient)null, (ITSAClient)null, 0, CryptoStandard.CMS);

            if (this.SignOnEveryPage)

                return;

            File.Delete(SourcePdfFileName);

        }

        public void SetFilePath()

        {

            StreamReader streamReader = new StreamReader("C:\\DigitalSignatureSettings\\FilesInfo.txt");

            string strSource;

            while ((strSource = streamReader.ReadLine()) != null)

            {

                if (strSource.Contains("OldFilePath:"))

                    this.sourcePath = this.getBetween(strSource, "OldFilePath:", "$");

                if (strSource.Contains("NewFilePath:"))

                    this.destinationPath = this.getBetween(strSource, "NewFilePath:", "$");

                if (strSource.Contains("REASON:"))

                    this.reason = this.getBetween(strSource, "REASON:", "$");

                if (strSource.Contains("LOCATION:"))

                    this.location = this.getBetween(strSource, "LOCATION:", "$");

                if (strSource.Contains("SIGNERNAME:"))

                    this.TokenCertificateName = this.getBetween(strSource, "SIGNERNAME:", "$");

                if (strSource.Contains("SignOnEveryPage:"))

                    this.SignOnEveryPage = Convert.ToBoolean(this.getBetween(strSource, "SignOnEveryPage:", "$"));

                ++this.num;

            }

        }

        public void splitdocument(string SourceFilePath)

        {

            string str = "C:\\DigitalSignatureSettings\\Split\\";
            PdfReader reader = new PdfReader(SourceFilePath);

            for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; ++pageNumber)

            {

                Document document = new Document();

                PdfCopy pdfCopy = new PdfCopy(document, (Stream)new FileStream(str + (object)pageNumber + ".pdf", FileMode.Create));

                document.Open();

                pdfCopy.AddPage(pdfCopy.GetImportedPage(reader, pageNumber));

                document.Close();

            }

        }

        public void mergedocument()

        {

            string[] files = Directory.GetFiles("C:\\DigitalSignatureSettings\\Split\\");

            string path1 = "C:\\DigitalSignatureSettings\\" + this.OutputFileName;

            Document document = new Document();

            PdfCopy pdfCopy = new PdfCopy(document, (Stream)new FileStream(path1, FileMode.Create));

            document.Open();

            for (int index = 0; index < files.Length; ++index)

            {

                PdfReader reader = new PdfReader(files[index]);

                for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; ++pageNumber)

                    pdfCopy.AddPage(pdfCopy.GetImportedPage(reader, pageNumber));

                reader.Close();

            }

            document.Close();

            foreach (string path2 in files)

                File.Delete(path2);

            File.Delete(this.sourcePath + this.OutputFileName);

            new Process() { StartInfo = { FileName = path1 } }.Start();

        }


       
    }
}