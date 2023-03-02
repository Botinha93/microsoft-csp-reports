param (
    [string]$type
 )
$outfile= "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN""  ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"">
<head>
<title>Report Clients</title>
<style>
#body{
    background-color: #f1f1f1;
    margin:0;
    overflow-x: hidden;
}
.container{
    width: 98%;
    margin-left:1%;
    font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;
}
.container table {
    font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 100%;
  }
  
.container table tbody, table th {
    border: 1px solid #ddd;
    padding: 8px;
  }
  
.container table tr:nth-child(even){background-color: #f9f9f9;}
  
.container table tr:hover {background-color: #ddd;}
  
.container table th {
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #4CAF50;
    color: white;
  }
  .collapsible {
    background-color: #777;
    color: white;
    cursor: pointer;
    padding: 18px;
    width: 100%;
    border: none;
    text-align: left;
    outline: none;
    font-size: 15px;
  }
  .active {
    background-color: #4CAF50;
  }
  .active, .collapsible:hover {
    background-color: #555;
  }

ul {
  list-style-type: none;
  margin: 0;
  padding: 0;
  overflow: hidden;
  background-color: #333;
  position: fixed;
  top: 0;
  width: 100%;
}

li {
  float: left;
}

li a {
  display: block;
  color: white;
  text-align: center;
  padding: 14px 16px;
  text-decoration: none;
  font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;
}

li a:hover:not(.active) {
  background-color: #111;
}
.frames{
    margin-top: 60px;
    height: calc(100vh - 60px);
    border: none;
  }

</style>
</head><body id=""body"">"

$navreports=$outfile +"
<ul>
  <li><a id=""sku_tab"" class=""active"" href=""javascript:doshow(`'sku`');"">Sku</a></li>
  <li><a id=""license_tab"" href=""javascript:doshow(`'license`');"">License</a></li>
  <li><a id=""users_tab"" href=""javascript:doshow(`'users`');"">Users</a></li>
  <li><a id=""reseller_tab"" href=""javascript:doshow(`'reseller`');"">Reseller</a></li>
  <li><a id=""billing_tab"" href=""javascript:doshow(`'billing`');"">Billing</a></li>
</ul>

<iframe id=""sku"" class=""frames"" style=""width:100%;display: block;"" src=""reports/report-sku.html"" /></iframe>
<iframe id=""license"" class=""frames"" style=""width:100%;display: none;"" src=""reports/report-license.html"" /></iframe>
<iframe id=""users"" class=""frames"" style=""width:100%;display: none;"" src=""reports/report-users.html"" /></iframe>
<iframe id=""reseller"" class=""frames"" style=""width:100%;display: none;"" src=""reports/report-reseller.html"" /></iframe>
<iframe id=""billing"" class=""frames"" style=""width:100%;display: none;"" src=""reports/report-billing.html"" /></iframe>
<script>
var lastone = document.getElementById(""sku"");
var lastonetab = document.getElementById(""sku_tab"");
function doshow(doc) {
    lastone.style.display = ""none"";
    var tochbange = document.getElementById(doc);
    tochbange.style.display = ""block"";
    lastone = tochbange;

    var tochbangetab = document.getElementById(doc.concat(""_tab""));
    tochbangetab.classList.add(""active""); 
    lastonetab.classList.remove(""active""); 
    lastonetab = tochbangetab;
}
        </script></body></html>"


    function report{
        $Customers = Get-PartnerCustomer
        foreach ($Customer in $Customers) {
            switch ($script:type) {
                "sku" {  
                    $CustomerReport = Get-PartnerCustomerSubscribedSku -CustomerId $Customer.CustomerId 2> Out-Null
                    if ($CustomerReport){
                        $CustomerReport = $CustomerReport | Select-Object productName,ActiveUnits,ConsumedUnits,WarningUnits,TotalUnits
                    }
                }
                "license" {  
                    $CustomerReport = Get-PartnerCustomerLicenseDeploymentInfo -CustomerId $Customer.CustomerId 2> Out-Null
                    if ($CustomerReport){
                        $CustomerReport = $CustomerReport | Select-Object productName,licensesDeployed,licensesSold,deploymentPercent
                    }
                }
                "users" {  
                    $CustomerReport = Get-PartnerCustomerUser -CustomerId $Customer.CustomerId 2> Out-Null
                    if ($CustomerReport){
                        $CustomerReport = $CustomerReport | Select-Object UserPrincipalName,DisplayName,PhoneNumber,State,UserDomainType,LastDirectorySyncTime
                    }
                }
                "reseller" {  
                    $CustomerReport = Get-PartnerIndirectReseller 2> Out-Null
                    if ($CustomerReport){
                        $CustomerReport = $CustomerReport | Select-Object Name,RelationshipType,State,Location
                    }
                }
                "billing" {  
                    $CustomerReport = Get-PartnerCustomerServiceCosts -BillingPeriod MostRecent -CustomerId $Customer.CustomerId 2> Out-Null
                    if ($CustomerReport){
                        $CustomerReport = $CustomerReport | Select-Object SubscriptionFriendlyName,StartDate,EndDate,ChargeType,UnitPrice,Quantity,Tax,PretaxTotal,AfterTaxTotal 
                    }
                }
                Default { 
                    Write-Host "Possible reports [sku, license, users, billing, reseller, all] all report will generate all reports and a navigation file called ""nav.html"""
                    break 
                }
            }
            $CustomerBody = $CustomerReport | ConvertTo-Html -Fragment | Out-String
            wiriteToOut -towirite $CustomerBody -title $Customer.Name
        }
        endReport
    }
    function resethml {
        $script:outfile= "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN""  ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">
        <html xmlns=""http://www.w3.org/1999/xhtml"">
        <head>
        <title>Report Clients</title>
        <style>
        #body{
            background-color: #f1f1f1;
            margin:0;
            overflow-x: hidden;
        }
        .container{
            width: 98%;
            margin-left:1%;
            font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;
        }
        .container table {
            font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;
            border-collapse: collapse;
            width: 100%;
          }
          
        .container table tbody, table th {
            border: 1px solid #ddd;
            padding: 8px;
          }
          
        .container table tr:nth-child(even){background-color: #f9f9f9;}
          
        .container table tr:hover {background-color: #ddd;}
          
        .container table th {
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: #4CAF50;
            color: white;
          }
          .collapsible {
            background-color: #777;
            color: white;
            cursor: pointer;
            padding: 18px;
            width: 100%;
            border: none;
            text-align: left;
            outline: none;
            font-size: 15px;
          }
          .active {
            background-color: #4CAF50;
          }
          .active, .collapsible:hover {
            background-color: #555;
          }
        
        ul {
          list-style-type: none;
          margin: 0;
          padding: 0;
          overflow: hidden;
          background-color: #333;
          position: fixed;
          top: 0;
          width: 100%;
        }
        
        li {
          float: left;
        }
        
        li a {
          display: block;
          color: white;
          text-align: center;
          padding: 14px 16px;
          text-decoration: none;
          font-family: ""Trebuchet MS"", Arial, Helvetica, sans-serif;
        }
        
        li a:hover:not(.active) {
          background-color: #111;
        }
        .frames{
            margin-top: 60px;
            height: calc(100vh - 60px);
            border: none;
          }
        
        </style>
        </head><body id=""body"">"
        
    }
    function endReport{
        $script:outfile+="<script>
        var coll = document.getElementsByClassName(""collapsible"");
        var i;
        
        for (i = 0; i < coll.length; i++) {
          coll[i].addEventListener(""click"", function() {
            this.classList.toggle(""active"");
            var content = this.nextElementSibling;
            if (content.style.display === ""table"") {
              content.style.display = ""none"";
            } else {
              content.style.display = ""table"";
            }
          });
          coll[i].nextElementSibling.style.display = ""table""
        }
        </script></body></html>"
        $script:outfile | Out-File "reports\report-$($script:type).html"
    }
    function wiriteToOut{
        param($towirite, $title)
        $script:outfile+="<div class=""container"" id=""$($title)""><button type=""button"" class=""collapsible"">$($title)</button>$($towirite)</div>"
    }

    if (Get-Module -ListAvailable -Name PartnerCenter) {
        Connect-PartnerCenter
        New-Item -ItemType Directory -Force -Path "reports"
        switch ($script:type) {
            "all" {  
                $script:type="sku"
                report
                resethml
                $script:type="license"
                report
                resethml
                $script:type="users"
                report
                resethml
                $script:type="billing"
                report
                resethml
                $script:type="reseller"
                report
                resethml
                $script:navreports | Out-File "nav.html"
            }
            Default { 
                report
            }
        }
    } 
    else {
        Install-Module -Name PartnerCenter -AllowClobber -Scope CurrentUser
    }
