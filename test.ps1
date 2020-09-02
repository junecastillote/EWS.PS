[CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [parameter(Mandatory,ParameterSetName='Default')]
        [parameter(Mandatory,ParameterSetName='DateFilter')]
        [ValidateNotNullOrEmpty()]
        $Token,

        [parameter(Mandatory,ParameterSetName='DateFilter')]
        [datetime]$StartDate,

        [parameter(Mandatory,ParameterSetName='DateFilter')]
        [datetime]$EndDate,

        [parameter(ParameterSetName='Default')]
        [switch]$All

    )

    $PSCmdlet.ParameterSetName