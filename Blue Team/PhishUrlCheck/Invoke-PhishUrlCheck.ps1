
function Get-PhishTankErrors
{
    $errors = New-Object System.Collections.Generic.List[System.Object]

    $erorrs.Add([pscustomobject] @{
        Timestamp = $req.response.meta.timestamp
        ErrorMessage = $req.response.results.errortext
        ServerId = $req.response.meta.serverid
        RequestId = $req.response.meta.requestid
    })
    return $errors
}

function New-HttpQueryString
{
    [CmdletBinding()]
    param 
    (
        [Parameter(Mandatory = $true)]
        [String]
        $Uri,
 
        [Parameter(Mandatory = $true)]
        [Hashtable]
        $QueryParameter
    )
 
    # Add System.Web
    Add-Type -AssemblyName System.Web
 
    # Create a http name value collection from an empty string
    $nvCollection = [System.Web.HttpUtility]::ParseQueryString([String]::Empty)
    
    foreach ($key in $QueryParameter.Keys)
    {
        $nvCollection.Add($key, $QueryParameter.$key)
    }
 
    # Build the uri
    $uriRequest = [System.UriBuilder]$uri
    $uriRequest.Query = $nvCollection.ToString()
 
    return $uriRequest.Uri.OriginalString
}


function Invoke-PhishUrlCheck
{
    #api documentation
    #https://www.phishtank.com/api_info.php
    param(
        [Parameter(Mandatory=$true)]
        [string]
        $key, 

        [Parameter(Mandatory=$false)]
        [string]
        $url
        )

    #or import the New-HttpQueryString funciton directly
    $uri = 'http://checkurl.phishtank.com/checkurl/index.php'
    $model = New-Object System.Collections.Generic.List[System.Object]

    #encoding to base64
    $urlBytes = [System.Text.Encoding]::UTF8.GetBytes($url)
    $urlId = [Convert]::ToBase64String($urlBytes)
    $urlId = $urlId.Replace('=','')

    $queryParams = @{
        url = $urlId
        format = 'json'
        app_key = $key
    }

    $headers = @{
        "User-Agent" = 'phishtank/MyPhishTankUsername'
    }

    $uriReq = New-HttpQueryString -Uri $uri -QueryParameter $queryParams

    $params = @{
        Uri = $uriReq
        Headers = $headers
        Method = 'POST'
    }

    try 
    {
        $req = Invoke-RestMethod @params
        $results = $req.response.results.url0
        $meta = $req.response.meta

        if($results.in_database -eq 'false')
        {
            $model.Add([pscustomobject] @{
                Timestamp = $meta.timestamp
                ServerId = $meta.serverid
                RequestId = $meta.requestid
                url = $results.url.'#cdata-section'
                InDataBase = $results.in_database
            })
            return $model
        }
        else
        {
            if($results.verified -eq 'true') #if the url has been verified on the site
            {
                $model.Add([pscustomobject] @{
                    Timestamp = $meta.timestamp
                    ServerId = $meta.serverid
                    RequestId = $meta.requestid
                    phish_id = $results.phish_id
                    url = $results.url.'#cdata-section'
                    in_database = $results.in_database
                    phish_detail_page = $results.phish_detail_page.'#cdata-section'
                    verified = $results.verified
                    verified_at = $results.verified_at
                    valid = $results.valid
                })
                return $model
            }
            else #phish is not verified on the site
            {
                $model.Add([pscustomobject] @{
                    Timestamp = $meta.timestamp
                    ServerId = $meta.serverid
                    RequestId = $meta.requestid
                    phish_id = $results.phish_id
                    url = $results.url.'#cdata-section'
                    in_database = $results.in_database
                    phish_detail_page = $results.phish_detail_page.'#cdata-section'
                    verified = $results.verified
                })
                return $model
            }
        }
    }
    catch #get any error that comes up and return object to the user
    {
        Get-PhishTankErrors
    }
}
