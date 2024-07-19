#Event ID 8321: A certificate validation operation took X milliseconds and has exceeded the execution time threshold
$RootCertFile="C:\SPRootCert.cer"
$SPRootCert = (Get-SPCertificateAuthority).RootCertificate
$SProotCert.Export("Cer") | Set-Content $RootCertFile -Encoding Byte

$RootCertFile="C:\SPRootCert.cer"
Import-Certificate -FilePath $RootCertFile -CertStoreLocation Cert:\LocalMachine\Root

#verify cert
#Start >> Run >> MMC >> File >> Add/Remove Snap-in 
#Click on Certificate >> Add >> Select computer account and click next >> select local computer and click finish.
#Expend Certificate >> Trusted Root Certification Authorities >> Certificate

#source https://www.sharepointdiary.com/2016/02/event-id-8321-certificate-validation-operation-took-x-milliseconds-exceeded-execution-time-threshold.html

