#requires -version 3.0

<#
   .Synoposis
   Contains cmdlets that get information about stock on the ASX and currency rates
    

#>


<#
    .Synopsis
    Gets the foreign currency exchange rate

    .Description
    Queries Yahoo Finance to get the currency rate between two currencies.

    .Parameter FromCurrency
    Three Letter Currency Code for the currency to convert from (source)

    .Parameter ToCurrency
    Three Letter Current Code for the currency to convert to (destination)

    .Example
    # Get the exchange rate from Australia Dollars to Japanese Yen
    Get-CurrencyRate -FromCurrency AUD -ToCurrency JPY

#>
Function Get-CurrencyRate 
{
    param (
     [string]$fromCurrency = $(throw "-FromCurrency is a required argument"),
     [string]$toCurrency = $(throw "-ToCurrency is a required argument")
    )

    try
    {
        $url = "http://finance.yahoo.com/d/quotes.csv?s=$($fromCurrency)$($toCurrency)=X&f=snl1";
        $result = invoke-RestMethod -Uri $url -Method Get;
        $items = $result.Split(",");
        return [float]$items[2];
    }
    catch 
    {
        Write-Error "Could not get the exchange rate";
    }
    return $null;
}



<#
    .Synopsis
    Creates an object for a stock holding

    .Description
    Creates an object for a stock holding, taking in the original buy order.

    .Parameter AsxCode
    The ASX code for a company

    .Parameter CompanyName
    The company name for display purposes
    
    .Parameter Quantity
    The original quantity of shares
    
    .Parameter BuyPrice
    The price at which the shares were acquired
    
#>
Function Get-HoldingInfo ([string]$asxCode, [string]$companyName, [int]$quantity, [single]$buyPrice)
{
    $holdingInfo = [PSCustomObject]@{
        AsxCode = $asxCode;
        FullName=$companyName;
        Quantity=$quantity;
        BuyPrice=$buyPrice;
    }
     $holdingInfo | Add-Member -NotePropertyName CurrentPrice -NotePropertyValue $(Get-Price -code $holdingInfo.AsxCode);
     $holdingInfo | Add-Member -NotePropertyName Value -NotePropertyValue $($holdingInfo.CurrentPrice * $holdingInfo.Quantity);
     $holdingInfo | Add-Member -NotePropertyName Cost -NotePropertyValue $($holdingInfo.BuyPrice * $holdingInfo.Quantity);
     $holdingInfo | Add-Member -NotePropertyName Profit -NotePropertyValue $($holdingInfo.Value - $holdingInfo.Cost);
     
     
     return $holdingInfo
}


<#
    .Synopsis
    Gets the delayed price of a stock listed on ASX.

    .Description
    Queries the Yahoo finance to get the last trade price (?confirm)
    for a stock


    .Parameter Code
    The three letter ASX code for a company

    .Parameter FormatAsMoney


    .Example
    # Get the price for Primary Health Care
    Get-Price PRY

    .Example
    # Get the price for Primary Health Care
    Get-Price -code PRY

#>
Function Get-Price 
{
    param ([string]$Code=$(throw "-Code is required argument. Code should be the three letter code used by ASX")
     ) 
    $retValue = $null;

    Try
    {
        $url = "http://finance.yahoo.com/d/quotes.csv?s=$($Code).AX&f=snl1d";
        $result = invoke-RestMethod -Uri $url -Method Get;

        $items = $result.Split(",");

       if( [single]::TryParse($items[2],[ref]$retValue) -eq $true) {
            $retValue =  [single]$items[2];
       }
   
    }
    Catch
    {
       Write-Error "Could not get a result for the supplied ASX code";
       

    }
    return $retValue;
    
}



<#
    .Synopsis
    Gets the list of prices of stocks listed on the ASX

    .Description
    For the supplied list of stocks, this function queries Yahoo Finance to find
    the last price.  The list of stocks use comma separated values of ASX code. 

    .Parameter StockList
    Comma separated list of ASX stock code to check

    .Example
    Get-StockPriceList ANZ, CBA, NAB, WBC

    .Example
    Get-StockPriceList -StockList ANZ, CBA, NAB, WBC

#>
Function Get-StockPriceList 
{
    param (
    [string]$stockList = $(throw "-StockList is a comma separated list of ASX codes to check.  It is a required argument")
    )
    $workItem = $stockList
    $option = [System.StringSplitOptions]::RemoveEmptyEntries;
    $stocks = $workItem.split()

    Write-Progress "Retreiving Prices..."

    $displayList = @();

    foreach ($stock in $stocks)
    {


        $price = Get-Price $stock


        $displayList += $(New-SimpleQuote -stock $stock -price $price)



    }
    $displayList | ft @{Expression={$_.Stock};Label="ASX Code"; width=10},
            @{Expression={$(if ($_.Price -lt 0.001) { "N/A" } else {
            ("{0:C2}" -f [single]$_.Price)})
            };Label="Price";}
}


<#
    .Synopsis
    creates an object containing a stock and a price

    .Description
    Creates a object with two properties for stock and price.  It is used
    to facilitate display.  

    This is for internal use only
    
    .Example
    New-SimpleQuote -stock CBA -price 70.00

#>
function New-SimpleQuote ([string]$stock, [single]$price)
{
    $quote = new-object -TypeName PSCustomObject
    $quote | Add-Member -type NoteProperty -name Stock -Value $stock.ToUpperInvariant()
    $quote | Add-Member -type NoteProperty -name Price -Value $price

    return $quote;
}


<#
    .Synopsis
    Displays information about a list of stock holdings

    .Description
    Displays the details about a stock.  Shows the current price of a stock,
    the original cost of the holidings and the current value of the holding.
    
    .Parameter Portfolio
    An array of stock holding information
     
#>
Function Show-Portfolio
{
    param (
        [PSCustomObject[]]$portfolio = $(throw "Need an array of holdings to be passed via the -portfolio parameter")
    )
    $portfolio | Sort-Object -Property FullName  |  ft @{Expression={ $_.FullName};Label="Company"},
@{Expression={ $_.Quantity}; Label="Shares"},
@{Expression={ ("{0:C3}" -f $_.BuyPrice)}; Label="Bought At"},
@{Expression={ ("{0:C2}" -f $_.Cost)}; Label="Cost"},
@{Expression={ ("{0:C3}" -f $_.CurrentPrice)}; Label="Current Price"},
@{Expression={ ("{0:C2}" -f $_.Value)}; Label="Value"},
@{Expression={ ("{0:C2}" -f $_.Profit) }; Label="Profit"} -AutoSize

}



<#
    .Sypnosis
    Displays information about stock prices based on a provided csv file

    .Description
    Displays the current price for stocks supplied in a csv file.  When
    the csv file contains heading of Quantity and Price it will calculate
    the current value and the purchase value.

    .Parameter FilePath
    Location for the scv file

    .Example
    Get-Portfolio "D:\local\scripts\powershell\finance\example\sunny_day.csv"

#>
function Get-Portfolio
{
    param (
        [string]$filePath =  $(throw "Requires a CSV file with the heading ASX for the code (Quantity and Price are optional headings) ")
    )

    $shortVersion = $false;
    $fullVersion = $false;

    # Check that there is file
    if($(Test-Path -Path $filePath -PathType Leaf -Include *.csv, *.txt ) -eq $true) {
        Write-Host $("File path {0} is valid" -f $filePath);
        

    } else {
        Write-Error $("File path {0} is not a valid *.txt or *.csv file" -f $filePath);    
        return;
    }

    # Check that it is sensisble
    $extraction = Import-Csv $filePath
    $stuff = $extraction | Get-Member;
    if(($stuff.Name -contains "asx") )
    {
        if(  ($stuff.Name -contains "price") -and 
         ($stuff.Name -contains "quantity") ) {
            FullVersion -content $extraction
            $fullVersion = $true;
        } else {
            ShortVersion -content $extraction
        }

    } else {
        Write-Error $("File path {0} does not have a header row with recognised headings" -f $filePath);    
    }
    return;
}






<#
    .Synopsis
    Display stock holdings plus a summary of worth and profit

    .Description
    Internal only


    .Example
    FullVersion -content $extraction
    #  where $extraction = Import-Csv $filePath

#>
function FullVersion ([PSObject]$content)
{

    $codeList = @();
    ForEach($item in $content)
    {
        $companyName = "";

        if( $($content | Get-Member).Name  -contains "company")
        {
            $companyName = $item.company;
        }

        [single]$thePrice = [convert]::ToSingle($item.price)

        $codeList += $(Get-HoldingInfo -asxCode $item.asx -companyName $companyName -quantity $item.quantity -buyPrice $thePrice);
    }

    Show-Portfolio -portfolio $codeList;

    $netWorth = ($myPortfolio | Measure-Object -Property Value -Sum).Sum
    $netProfit = ($myPortfolio | Measure-Object -Property Profit -Sum).Sum



    Write-Output $("Portfolio Value is: {0:C2} Profit ({1:C2})" -f $netWorth, $netProfit)

    return;
}


<#
    .Sypnosis
    Displays the current prices for the supplied content

    .Description
    Internal only function

     .Example
    ShortVersion -content $extraction
    #  where $extraction = Import-Csv $filePath

#>
function ShortVersion ([PSObject]$content)
{
    $codeList = @();
    ForEach($item in $content)
    {
        $codeList += $item.Asx;
    }

    Get-StockPriceList -stockList $codeList;
    return;
}



Export-ModuleMember -Function Get-CurrencyRate
Export-ModuleMember -Function Get-Price 
Export-ModuleMember -Function Get-StockPriceList
Export-ModuleMember -Function Get-HoldingInfo
Export-ModuleMember -Function Get-Portfolio
Export-ModuleMember -Function Show-Portfolio