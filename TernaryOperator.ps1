function ??{
    PARAM(
    [parameter(ValueFromPipeline, Mandatory)]
    [bool]$bool,
    [parameter(Position = 0,Mandatory=$true)]
    $trueVal,
    [parameter(Position = 1,Mandatory=$true)]
    [ValidatePattern(":")]$s,
    [parameter(Position = 2,Mandatory=$true)]
    $flaseVal
    )if($bool){$trueValue} else {$flaseValue}
}
