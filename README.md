# Endpoint-DLP-Microsoft
## DLP Test WebSites
> -   Test DLP Site1 https://testdlp.thehpc.in/ 
> - Test DLP Site2 https://testdlp.azurewebsites.net/

<br>

## Here is Powershel Commadlets Samples.

#### <u> # Run below command to get all the DLP Policy </u> 

```sh
 Get-DlpCompliancePolicy 

```
<u> Output </u>

 ![alt text](https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft/blob/main/img/output1.png?raw=true)

#### <u> # Get Policy Details for a specific Policy </u>
```sh
 Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext"

 or

 Get-DlpCompliancePolicy "MyDLP-Test01-Confidentialtext" | Where-Object -Property EndpointDlpLocation -NE ""|Select-Object Priority,DisplayName,Mode,EndpointDlpLocation,EndpointDlpLocationException


```
<u> Output </u>
 
![alt text](https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft/blob/main/img/Output2.png?raw=true)
 

<br>
<br>
<br>
<br>


Developed by Rajeshcbsa
https://github.com/Rajeshcbsa/Endpoint-DLP-Microsoft
