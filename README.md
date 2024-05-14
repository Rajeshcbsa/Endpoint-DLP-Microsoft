# Endpoint-DLP-Microsoft
Microsoft Endpoint DLP

Test DLP Site https://testdlp.thehpc.in/



## Get All Dlp Policy ✨
```sh
Get-DlpCompliancePolicy 

```
Output
 
![alt text](https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft/blob/main/img/output1.png?raw=true)

## Get Policy Details for a specific Policy
```sh
Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext"

or

Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext" | Where-Object -Property EndpointDlpLocation -NE ""|Select-Object Priority,DisplayName,Mode,EndpointDlpLocation,EndpointDlpLocationException


```
Output
 
![alt text](https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft/blob/main/img/Output2.png?raw=true)
 

Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext" | Where-Object -Property EndpointDlpLocation -NE ""|Select-Object Priority,DisplayName,Mode,EndpointDlpLocation,EndpointDlpLocationException

Developed by Rajeshcbsa
https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft
