$passphrase = ConvertTo-SecureString -String “2FJlsXghEas5vdJJKEXXwWFab” -asPlainText -Force
Set-SPPassPhrase -PassPhrase $passphrase -Confirm
