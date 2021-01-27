function WinApi{
    <#
    .SYNOPSIS
        Вызывает WinApi  функции
        (Calls WinApi functions)
    .EXAMPLE
        Examples are on the link
    .NOTES
        Author: Fors1k
    .LINK
        https://www.cyberforum.ru/post14941731.html
    #>
    PARAM(
        [Parameter(Position = 0, Mandatory = $True )]
        [String]
        $method,
        [Parameter(Position = 1, Mandatory = $True )]
        [Object[]]
        $params,
        [Parameter(Position = 2, Mandatory = $False)]
        [Type]
        $return  = [Boolean],
        [Parameter(Position = 3, Mandatory = $False)]
        [String]
        $dll     = 'User32.dll',
        [Parameter(Position = 4, Mandatory = $False)]
        [Int]
        $charSet = 4
    )
    BEGIN{
        [Type[]]$pTypes = $params|ForEach{
            if($_-is[ref]){$_.Value.GetType().MakeByRefType()}
            elseif($_-is[scriptblock]){
                [Delegate]::CreateDelegate([func[[type[]],type]],
                [Linq.Expressions.Expression].Assembly.GetType(
                'System.Linq.Expressions.Compiler.DelegateHelpers'
                ).GetMethod('MakeNewCustomDelegate',
                [Reflection.BindingFlags]'NonPublic, Static') 
                ).Invoke(($_.ast.ParamBlock.Parameters.StaticType+
                $_.Attributes.type.type))
            }else{$_.GetType()}
        }
    }
    PROCESS{
        ($w32=[Reflection.Emit.AssemblyBuilder]::
        DefineDynamicAssembly('w32A','Run').
        DefineDynamicModule('w32M').DefineType('w32T',
        "Public,BeforeFieldInit")).DefineMethod(
        $method,'Public,HideBySig,Static,PinvokeImpl',
        $return,($pTypes)).SetCustomAttribute(
        [Reflection.Emit.CustomAttributeBuilder]::new(
        ($DI=[Runtime.InteropServices.DllImportAttribute]).
        GetConstructor([string]),$dll,$DI.GetField('CharSet'),
        @{[char]'W'=-1;;[char]'A'=-2}[$method[-1]]+$CharSet))
    }
    END{$w32.CreateType()::$method.Invoke($Params)}
}
