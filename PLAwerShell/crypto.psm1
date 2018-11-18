#-------------------------------------------------------
#region Crypto Functions
# Credit: SDWheeler

#-------------------------------------------------------
function new-password {
    <#
    .Description
        Returns a password
    #>
    param($length = 16, $numspecials = 6)
    [System.Web.Security.Membership]::GeneratePassword($length, $numspecials)
}
#-------------------------------------------------------
function show-certificate {
    param([string[]]$certificate,
        [switch]$multiline)

    foreach ($certpath in $certificate) {
        $cert = get-item $certpath
        $cert | Select-Object -Property Subject, DNSNameList, NotBefore, NotAfter, Issuer, PolicyId, Archived, FriendlyName, SerialNumber, Thumbprint, HasPrivateKey,
        @{ N = 'SignatureAlgorithm'; E = { $_.SignatureAlgorithm.FriendlyName } },
        @{ N = 'CertTemplateInfo'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Certificate Template Information'}).Format($multiline)} },
        @{ N = 'KeyUsage'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Key Usage'}).Format($multiline)} },
        @{ N = 'EnhKeyUsage'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Enhanced Key Usage'}).Format($multiline)} },
        @{ N = 'AppPolicies'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Application Policies'}).Format($multiline)} },
        @{ N = 'SubjectKeyId'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Subject Key Identifier'}).Format($multiline)} },
        @{ N = 'AuthorityKeyId'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Authority Key Identifier'}).Format($multiline)} },
        @{ N = 'CRLDistPoints'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'CRL Distribution Points'}).Format($multiline)} },
        @{ N = 'AuthorityInfoAccess'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Authority Information Access'}).Format($multiline)} },
        @{ N = 'SubjectAltName'; E = { ($_.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Subject Alternative Name'}).Format($multiline)} },
        @{ N = 'BasicConstraints'; E = { ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq 'Basic Constraints'}).Format($multiline)} }
    }
}
#endregion
#-------------------------------------------------------
