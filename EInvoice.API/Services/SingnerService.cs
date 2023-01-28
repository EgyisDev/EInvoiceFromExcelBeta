using EInvoice.Common.Configurations;
using Microsoft.Extensions.Options;
using Net.Pkcs11Interop.Common;
using Net.Pkcs11Interop.HighLevelAPI;
using Net.Pkcs11Interop.HighLevelAPI.Factories;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Ess;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace EInvoice.Services;

public class SignerService : ISignerService
{
    private readonly string DllLibPath = "eps2003csp11.dll";
    private string TokenCertificate = "Egypt Trust Sealing CA";

    private readonly IPkcs11LibraryFactory _factory;

    private ApplicationConfiguration _applicationConfiguration;
    private readonly EInvoicingConfiguration _eInvoicingConfiguration;

    public SignerService(IOptions<ApplicationConfiguration> applicationConfiguration, IPkcs11LibraryFactory factory, IOptions<EInvoicingConfiguration> eInvoicingConfiguration)
    {
        _factory = factory;
        _eInvoicingConfiguration = eInvoicingConfiguration.Value;
        _applicationConfiguration = applicationConfiguration.Value;
    }

    public SignerService(EInvoicingConfiguration eInvoicingConfiguration)
    {
        _eInvoicingConfiguration = eInvoicingConfiguration;
    }

    public string SignWithCMS(string serializedJson)
    {
        var data = Encoding.UTF8.GetBytes(serializedJson);

        var factories = new Pkcs11InteropFactories();

        using var pkcs11Library = _factory.LoadPkcs11Library(factories, DllLibPath, AppType.MultiThreaded);

        var slot = pkcs11Library.GetSlotList(SlotsType.WithTokenPresent).FirstOrDefault();

        if (slot is null)
        {
            throw new Exception("No slots found");
        }

        using var session = slot.OpenSession(SessionType.ReadWrite);

        session.Login(CKU.CKU_USER, Encoding.UTF8.GetBytes(_eInvoicingConfiguration.TokenPin));

        var certificateSearchAttributes = new List<IObjectAttribute>()
        {
            session.Factories.ObjectAttributeFactory.Create(CKA.CKA_CLASS, CKO.CKO_CERTIFICATE),
            session.Factories.ObjectAttributeFactory.Create(CKA.CKA_TOKEN, true),
            session.Factories.ObjectAttributeFactory.Create(CKA.CKA_CERTIFICATE_TYPE, CKC.CKC_X_509)
        };

        var certificate = session.FindAllObjects(certificateSearchAttributes).FirstOrDefault();

        if (certificate is null)
        {
            throw new Exception("Certificate not found");
        }

        var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.MaxAllowed);

        // find cert by thumbprint
        var foundCerts = store.Certificates.Find(X509FindType.FindByIssuerName, TokenCertificate, false);

        //var foundCerts = store.Certificates.Find(X509FindType.FindBySerialNumber, "520456a1a5d5204b0deabf8bbafb65e4", true);


        if (foundCerts.Count == 0)
            throw new Exception("No device detected");

        var certForSigning = foundCerts[0];
        store.Close();


        var content = new ContentInfo(new Oid("1.2.840.113549.1.7.5"), data);


        var cms = new SignedCms(content, true);

        var bouncyCertificate = new EssCertIDv2(new Org.BouncyCastle.Asn1.X509.AlgorithmIdentifier(new DerObjectIdentifier("1.2.840.113549.1.9.16.2.47")), this.HashBytes(certForSigning.RawData));

        var signerCertificateV2 = new SigningCertificateV2(new EssCertIDv2[] { bouncyCertificate });


        var signer = new CmsSigner(certForSigning);

        signer.DigestAlgorithm = new Oid("2.16.840.1.101.3.4.2.1");



        signer.SignedAttributes.Add(new Pkcs9SigningTime(DateTime.UtcNow));
        signer.SignedAttributes.Add(new AsnEncodedData(new Oid("1.2.840.113549.1.9.16.2.47"), signerCertificateV2.GetEncoded()));


        cms.ComputeSignature(signer);

        var output = cms.Encode();

        return Convert.ToBase64String(output);
    }

    private byte[] HashBytes(byte[] input)
    {
        using var sha = SHA256.Create();
        var output = sha.ComputeHash(input);
        return output;
    }

    public void ListCertificates()
    {

        var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.MaxAllowed);
        var collection = store.Certificates;
        var foundCollection = collection.Find(X509FindType.FindBySerialNumber, "27facef6632e6c8e4e086214cf6c9be6", true);

        foreach (var x509 in foundCollection)
        {
            try
            {
                var rawData = x509.RawData;
                Console.WriteLine("Content Type: {0}{1}", X509Certificate2.GetCertContentType(rawData), Environment.NewLine);
                Console.WriteLine("Friendly Name: {0}{1}", x509.FriendlyName, Environment.NewLine);
                Console.WriteLine("Certificate Verified?: {0}{1}", x509.Verify(), Environment.NewLine);
                Console.WriteLine("Simple Name: {0}{1}", x509.GetNameInfo(X509NameType.SimpleName, true), Environment.NewLine);
                Console.WriteLine("Signature Algorithm: {0}{1}", x509.SignatureAlgorithm.FriendlyName, Environment.NewLine);
                Console.WriteLine("Public Key: {0}{1}", x509.PublicKey.Key.ToXmlString(false), Environment.NewLine);
                Console.WriteLine("Certificate Archived?: {0}{1}", x509.Archived, Environment.NewLine);
                Console.WriteLine("Length of Raw Data: {0}{1}", x509.RawData.Length, Environment.NewLine);
                x509.Reset();
            }
            catch (CryptographicException ex)
            {
                Console.WriteLine("Information could not be written out for this certificate.");
                throw ex;
            }
        }
        store.Close();
    }
}



