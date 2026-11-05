<#
.SYNOPSIS
Creates a self-signed certificate for the Documents Entra app registration.

.DESCRIPTION
Writes two files:
  <name>.cer  public key; upload it to the app registration (Certificates & secrets > Certificates).
  <name>.pfx  private key protected by -Password; upload it to the certificate HTTP connection
              (HTTP with Microsoft Entra ID (preauthorized), Certificate authentication).
The key is held in a temporary CryptoAPI key container that is deleted when the script finishes;
the certificate is never added to a Windows certificate store. Files are named
<common name>-<date>-<thumbprint>, so a rotation never overwrites an earlier certificate.

The default output folder is .local/certificates, which git ignores. Keep the .pfx out of this
repository and delete it once it is uploaded to the connection.

.EXAMPLE
./scripts/new-app-certificate.ps1
Prompts for the .pfx password and writes a 12-month certificate to .local/certificates.

.EXAMPLE
./scripts/new-app-certificate.ps1 -Subject 'Contoso Documents worker' -ValidityMonths 24 -OutputDirectory C:\secure\certs
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][SecureString]$Password,
    [string]$Subject = 'CN=Ascentix Documents worker',
    # Entra sets no limit on a self-signed certificate's lifetime; this is the rotation cadence.
    [ValidateRange(1, [int]::MaxValue)][int]$ValidityMonths = 12,
    [ValidateSet(2048, 3072, 4096)][int]$KeySize = 2048,
    [string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
# A plain name such as 'contoso-documents' becomes 'CN=contoso-documents'.
if ($Subject -notmatch '=') { $Subject = "CN=$Subject" }
# Files are named after the common name; characters Windows forbids in file names become '-'.
if ($Subject -notmatch '(?:^|,)\s*CN=([^,]+)') { throw "Subject '$Subject' has no CN (common name)." }
$commonName = $Matches[1].Trim()
foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $commonName = $commonName.Replace($c, '-') }
$productRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $productRoot '.local/certificates' }
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($OutputDirectory) | Out-Null

# The connector reads the .pfx key through legacy CryptoAPI and fails with "Invalid provider type
# specified" on a CNG (Key Storage Provider) key, so the key lives in the RSA/AES CSP (type 24).
$csp = [Security.Cryptography.CspParameters]::new(24, 'Microsoft Enhanced RSA and AES Cryptographic Provider')
$csp.KeyContainerName = 'documents-app-' + [Guid]::NewGuid().ToString('N')
$key = [Security.Cryptography.RSACryptoServiceProvider]::new($KeySize, $csp)
$key.PersistKeyInCsp = $false
$request = [Security.Cryptography.X509Certificates.CertificateRequest]::new(
    [Security.Cryptography.X509Certificates.X500DistinguishedName]::new($Subject), $key,
    [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
$request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $true))
$request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
    [Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature, $true))
$clientAuth = [Security.Cryptography.OidCollection]::new()
$clientAuth.Add([Security.Cryptography.Oid]::new('1.3.6.1.5.5.7.3.2')) | Out-Null
$request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($clientAuth, $false))
$request.CertificateExtensions.Add([Security.Cryptography.X509Certificates.X509SubjectKeyIdentifierExtension]::new($request.PublicKey, $false))

$notBefore = [DateTimeOffset]::UtcNow.AddMinutes(-5)
$certificate = $request.CreateSelfSigned($notBefore, $notBefore.AddMonths($ValidityMonths))
$thumbprint = $certificate.Thumbprint
try {
    $name = '{0}-{1:yyyyMMdd}-{2}' -f $commonName, $notBefore, $thumbprint.Substring(0, 8).ToLowerInvariant()
    $cerPath = Join-Path $OutputDirectory "$name.cer"
    $pfxPath = Join-Path $OutputDirectory "$name.pfx"
    $plainPassword = [Net.NetworkCredential]::new('', $Password).Password
    [IO.File]::WriteAllBytes($cerPath, $certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    [IO.File]::WriteAllBytes($pfxPath, $certificate.Export([Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $plainPassword))
}
finally {
    $plainPassword = $null
    $certificate.Dispose()
    $key.Dispose()
}

[pscustomobject]@{
    Subject    = $Subject
    Thumbprint = $thumbprint
    NotBefore  = $notBefore.UtcDateTime
    NotAfter   = $notBefore.AddMonths($ValidityMonths).UtcDateTime
    Cer        = $cerPath
    Pfx        = $pfxPath
}
Write-Host "Next: upload $name.cer to the app registration, then $name.pfx with its password to the certificate HTTP connection. Delete the .pfx after the connection is saved."
