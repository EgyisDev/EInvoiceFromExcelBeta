

using System.Security.Cryptography.X509Certificates;

namespace EInvoice.Services;

public class SignerServiceV2 : ISignerService
{
    public string SignWithCMS(string serializedJson)
    {
        var store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
        store.Open(OpenFlags.MaxAllowed);

        // Find the certificate by subject name
        var certificates = store.Certificates.Find(X509FindType.FindBySerialNumber, "27facef6632e6c8e4e086214cf6c9be6", false);
        
        if (certificates.Count > 0)
        {
            X509Certificate2 certificate = certificates[0];
            Console.WriteLine("Subject name: {0}", certificate.SubjectName.Name);
        }

        store.Close();

        throw new NotImplementedException();
    }
}

