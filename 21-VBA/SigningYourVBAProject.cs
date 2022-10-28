using OfficeOpenXml;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    class SigningYourVBAProject
    {
        /// <summary>
        /// Opens the Battleships sample and sign it with the certificate from the pfx file.
        /// </summary>
        public static void Run()
        {
            //Load our test certificate from the pfx file. 
            //In a real production environment, make sure you store your certificate in a secure way.
            var cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(FileUtil.GetRootDirectory() + "\\21-VBA\\SampleCertificate.pfx", "EPPlus");     
            
            //Open the workbook created in the previous sample.
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("21.3-CreateABattleShipsGameVba.xlsm")))
            {
                var signature = p.Workbook.VbaProject.Signature;

                //The only thing you need to do to sign your project is to set the signatures 'Certificate' property with your code-signing certificate.
                //Your certificate must have access to the private key to sign the project.
                signature.Certificate = cert;

                //If the file is unsigned, EPPlus will by default create all three signatures - Legacy, Agile and V3.
                //You can use the property 'CreateSignatureOnSave' to decide which signature version you want to create on saving the workbook.
                //For example 'signature.LegacySignature.CreateSignatureOnSave = false' to remove the legacy signature.

                //You can also set the hash algorithm for each signature version.
                //Excel and EPPlus default is MD5 for the legacy signature and SHA1 for the Agile and V3 signature.
                //We want to change it to SHA256 to get better and more modern hash algorithm.
                signature.LegacySignature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA256; 
                signature.AgileSignature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA256;  
                signature.V3Signature.HashAlgorithm = VbaSignatureHashAlgorithm.SHA256;

                p.SaveAs(FileUtil.GetFileInfo("21.4-Signed-CreateABattleShipsGameVba.xlsm"));
            }            
        }

    }
}
