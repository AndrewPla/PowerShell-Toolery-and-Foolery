function wootrss {
  param(
    [ValidateSet('accessories','computers','electronics','home','kids','sellout','shirt','sport','tools','wine','www', ignorecase=$true)]
    [string]$site,
    [switch]$notable
  )
  $deals = @()
  $url = 'http://api.woot.com/1/sales/current.rss'
  $woots = Invoke-RestMethod $url | Where-Object {$_.link.startswith("https://$site") }

  foreach ($woot in $woots) {
    $wootoff = ''
    if ($woot.wootoff -eq 'true') {$wootoff = 'woot!'}
    $props = [ordered]@{site=($woot.link -split '\.')[0] -replace 'https://','';
      title=$woot.title;
      price=$woot.pricerange;
      '%sold'=[double]($woot.soldoutpercentage) * 100;
      wootoff=$wootoff;
      condition=$woot.condition;
    }
    $deals += new-object -type PSObject -prop $props
  }
  if ($notable) {
    $deals
  } else {
    $deals | Format-Table -AutoSize
  }
}
function woot {
  param([switch]$notable=$false)
  $apikey = Get-PSFConfigValue -fullname plawershell.wootKey
  $url = 'https://api.woot.com/2/events.json?eventType=Daily&key={0}' -f  $apikey
  $daily = invoke-restmethod $url
  $url = 'https://api.woot.com/2/events.json?eventType=WootOff&key={0}' -f  $apikey
  $daily += invoke-restmethod $url
  $results = $daily | sort site |
  Select-Object `
  @{l='site';e={($_.site -split '\.')[0]}},
  type,
  @{l='title';e={$_.offers.Title}},
  @{l='Price';e={$_.offers.items.SalePrice | Sort-Object | Select-Object -first 1 -Last 1}},
  @{l='%Sold';e={100 - $_.offers.PercentageRemaining}},
  @{l='Condition';e={$_.offers.items.Attributes | Where-Object Key -eq 'Condition' | Select-Object -ExpandProperty Value -First 1}}
  if ($notable) {$results} else {$results | ft -AutoSize}
}