Param(
    [Parameter(Mandatory=$True)] 
    [string]$projectFile = "",
    [Parameter(Mandatory=$False)] 
    [int]$updateIntervalInDays = 0
)

if ($env:OfflinePublishUrl -eq $null)
{
    $(throw "OfflinePublishUrl is a required variable")
}

if ($env:CertificateThumbPrint -eq $null)
{
    $(throw "CertificateThumbPrint is a required variable")
}

if ($env:TimestampUrl -eq $null)
{
    $(throw "TimestampUrl is a required variable")
}

function Edit-XmlNodes {
param (
    [xml] $doc = $(throw "doc is a required parameter"),
    [string] $xpath = $(throw "xpath is a required parameter"),
    [string] $value = $(throw "value is a required parameter"),
    [string] $namespace,
    [string] $namespacePrefix,
    [bool] $condition = $true
)   

    [System.Xml.XmlNamespaceManager] $nsmgr;
    if ($namespace -ne $null -and $namespacePrefix -ne $null)
    {
        if ($namespacePrefix -eq "default")
        { $namespacePrefix = $null}

        [System.Xml.XmlNamespaceManager] $nsmgr = $doc.NameTable;
        $nsmgr.AddNamespace("",$namespace);
    }
    
    if ($condition -eq $true) {
        if ($nsmgr -eq $null)
        {
            $nodes = $doc.SelectNodes($xpath)
        }
        else 
        {
            $nodes = $doc.SelectNodes($xpath,$nsmgr)
        }
         
        foreach ($node in $nodes) {
            if ($node -ne $null) {
                if ($node.NodeType -eq "Element") {
                    $node.InnerXml = $value
                }
                else {
                    $node.Value = $value
                }
            }
        }
    }
}

# Workaround for Xpath issues
$content = Get-Content -Path $projectFile
$content = $content.Replace('xmlns','xmlns:n');
Set-Content -Path $projectFile -Value $content

$xml = new-object System.Xml.XmlDocument;
$xml.Load($projectFile)

$elements = $xml.SelectNodes("//PublishUrl");
$elements | ForEach-Object {$_.InnerText = "$env:OfflinePublishUrl"};
Write-Host "Updated $($elements.Count) PublishUrl elements with value $env:OfflinePublishUrl";

#ToDo: is not set in orginal build process - needs to be checked
$elements = $xml.SelectNodes("//InstallUrl");
$elements | ForEach-Object {$_.InnerText = "$env:OfflinePublishUrl"};
Write-Host "Updated $($elements.Count) InstallUrl elements with value $env:OfflinePublishUrl";

#ToDo: is not set in orginal build process - needs to be checked
$elements = $xml.SelectNodes("//UpdateUrl");
$elements | ForEach-Object {$_.InnerText = "$env:OfflinePublishUrl"};
Write-Host "Updated $($elements.Count) UpdateUrl elements with value $env:OfflinePublishUrl";

$elements = $xml.SelectNodes("//PropertyGroup/Install");
$elements | ForEach-Object {$_.InnerText = "true"};
Write-Host "Updated $($elements.Count) Install elements with value true";

$elements = $xml.SelectNodes("//UpdateEnabled");
$elements | ForEach-Object {$_.InnerText = "false"};
Write-Host "Updated $($elements.Count) UpdateEnabled elements with value false";

$elements = $xml.SelectNodes("//IsWebBootstrapper");
$elements | ForEach-Object {$_.InnerText = "false"};
Write-Host "Updated $($elements.Count) IsWebBootstrapper elements with value false";

$elements = $xml.SelectNodes("//UpdateInterval");
$elements | ForEach-Object {$_.InnerText = "$updateIntervalInDays"};
Write-Host "Updated $($elements.Count) UpdateInterval elements with value $updateIntervalInDays";

$elements = $xml.SelectNodes("//UpdateIntervalUnits");
$elements | ForEach-Object {$_.InnerText = "Days"};
Write-Host "Updated $($elements.Count) UpdateIntervalUnits elements with value Days";

$elements = $xml.SelectNodes("//ManifestTimestampUrl");
$elements | ForEach-Object {$_.InnerText = "$env:TimestampUrl"};
Write-Host "Updated $($elements.Count) ManifestTimestampUrl elements with value $env:TimestampUrl";

$elements = $xml.SelectNodes("//ManifestCertificateThumbprint");
$elements | ForEach-Object {$_.InnerText = "$env:CertificateThumbPrint"};
Write-Host "Updated $($elements.Count) ManifestCertificateThumbprint elements with value $env:CertificateThumbPrint";

$elements = $xml.SelectNodes("//ApplicationVersion");
$elements | ForEach-Object {$_.InnerText = "$env:BUILD_BUILDNUMBER"};
Write-Host "Updated $($elements.Count) ApplicationVersion elements with value $env:BUILD_BUILDNUMBER";

$xml.Save($projectFile);

$content = Get-Content -Path $projectFile
$content = $content.Replace('xmlns:n','xmlns');
Set-Content -Path $projectFile -Value $content