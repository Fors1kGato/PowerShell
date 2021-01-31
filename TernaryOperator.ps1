function ??{
    PARAM(
    [parameter(ValueFromPipeline, Mandatory)]
    [Boolean]$bool,
    [parameter(Position = 0,Mandatory=$true)]
    $trueValue,
    [parameter(Position = 1,Mandatory=$true)]
    [ValidatePattern(":")]$s,
    [parameter(Position = 2,Mandatory=$true)]
    $flaseValue
    )if($bool){$trueValue} else {$flaseValue}
}
