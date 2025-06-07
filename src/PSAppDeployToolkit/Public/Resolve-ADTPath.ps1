#-----------------------------------------------------------------------------
#
# MARK: Resolve-ADTPath
#
#-----------------------------------------------------------------------------

function Resolve-ADTPath {
    <#
    .SYNOPSIS
        Resolves a given path to its full path.

    .DESCRIPTION
        Resolves a given path to its fully qualified PowerShell paths, or alternatively to the provider path.
        It can take any object containing a PSPath. These are usually objects retrieved by cmdlets like Get-Item or Get-ChildItem.
        This function imitates the behavior of the PowerShell internal LocationGlobber.
        If no provider specific parameters are specified, it behaves equivalent to the Resolve-Path cmdlet.

    .PARAMETER Path
        The path to resolve.
        If the provider supports wildcards, the path can contain wildcard characters.

    .PARAMETER LiteralPath
        The literal path to resolve.

    .PARAMETER Filter
        A filter to qualify the paths.

    .PARAMETER Exclude
        An array of patterns to exclude.

    .PARAMETER Include
        An array of patterns to include.

    .PARAMETER Force
        Forces the resolution of hidden or system items.

    .PARAMETER PathType
        Limits the output to the given path type.
        By default, any path type is accepted.
        Options: Any, Container, Leaf

    .PARAMETER IncludeNonExistent
        Includes non-existent paths in the output.

    .PARAMETER AsProviderPath
        Returns the provider internal path instead of the fully qualified PowerShell path.
        This is useful for working with .NET objects that do not support provider paths.

    .PARAMETER ProviderName
        Will validate the given path against the specified provider.
        Specifying this parameter will enable to translate provider paths to their fully qualified PowerShell paths.
        It is recommended to specify the provider name to limit the output to a specific provider.
        The parameter must be a valid provider name from Get-PSProvider.
    .PARAMETER UnboundArguments
        This will be ignored. Its solely purpose is to make splatting PSBoundParameters easier.

    .INPUTS
        System.Object[]

        Resolve-ADTPath will accept any object that includes a PSPath property

    .OUTPUTS
        System.String[]

        Returns an array of strings representing the resolved paths.

    .EXAMPLE
        Resolve-ADTPath -Path "C:\Windows\System32\*" -Filter '*.exe' -Include 'cmd*'

        Returns all executable files in the System32 directory that start with 'cmd'.

    .EXAMPLE
        Resolve-ADTPath @PSBoundParameters -IncludeNonExistent

        Resolves parameters passed to the parent function and includes non-existent paths in the output, if they can be resolved.

    .EXAMPLE
        Get-Item -Path "HKLM:\Software" | Resolve-ADTPath -AsProviderPath

        Returns the "HKEY_LOCAL_MACHINE\Software" provider path for the given registry key.

    .LINK
        https://psappdeploytoolkit.com/docs/reference/functions/Resolve-ADTPath

    .LINK
        https://github.com/PowerShell/PowerShell/blob/master/src/System.Management.Automation/namespaces/LocationGlobber.cs

    #>
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Variables are used in the process block')]
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    [OutputType([string[]])]
    param (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'Path')]
        [SupportsWildcards()][ValidateNotNullOrEmpty()]
        [string[]]$Path,

        [Parameter(Mandatory = $false, ParameterSetName = 'LiteralPath', ValueFromPipelineByPropertyName = $true)]
        [Alias('PSPath')]
        [ValidateNotNullOrEmpty()]
        [string[]]$LiteralPath,

        [Parameter(Mandatory = $false)]
        [SupportsWildcards()]
        [string]$Filter,

        [Parameter(Mandatory = $false)]
        [SupportsWildcards()]
        [string[]]$Exclude,

        [Parameter(Mandatory = $false)]
        [SupportsWildcards()]
        [string[]]$Include,

        [Parameter(Mandatory = $false)]
        [switch]$Force,

        [Parameter(Mandatory = $false)]
        [Microsoft.PowerShell.Commands.TestPathType]$PathType = 'Any',

        [Parameter(Mandatory = $false)]
        [switch]$IncludeNonExistent,

        [Parameter(Mandatory = $false)]
        [switch]$AsProviderPath,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
                try
                {
                    $null = $ExecutionContext.SessionState.Provider.GetOne($_)
                }
                catch
                {
                    $PSCmdlet.ThrowTerminatingError((New-ADTValidateScriptErrorRecord -ParameterName ProviderName -ProvidedValue $_ -ExceptionMessage "The provided `-ProviderName` parameter does not match any installed PowerShell provider."))
                }
            })]
        [string]$ProviderName,

        [Parameter(Mandatory = $false, DontShow = $true, ValueFromRemainingArguments)]
        [AllowNull()][AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$UnboundArguments
    )

    begin
    {
        Initialize-ADTFunction -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    }

    process
    {
        try
        {
            $isLiteralPath = $PSBoundParameters.ContainsKey('LiteralPath')

            $(if ($isLiteralPath) { $LiteralPath } else { $Path }) |
            & {
                process
                {
                    # Determine the provider and the path. FileSystem is the default provider if no provider is specified.
                    $providerInfo = $null
                    $providerPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_, [ref]$providerInfo, [ref]$null)

                    # Validate the provider
                    if ($ProviderName)
                    {
                        $desiredProvider = $ExecutionContext.SessionState.Provider.GetOne($ProviderName)
                        # If the path is not provider qualified, just use the desired provider.
                        if (-not $ExecutionContext.SessionState.Path.IsProviderQualified($_))
                        {
                            $providerInfo = $desiredProvider
                            $providerPath = $_
                        }
                        # if the path is provider qualified, check if it matches the desired provider.
                        elseif (-not $providerInfo.Equals($desiredProvider))
                        {
                            $naerParams = @{
                                Exception = [System.Management.Automation.ProviderInvocationException]::new("The given path '$_' is not valid for the specified provider '$($providerInfo.Name)'.")
                                Category = [System.Management.Automation.ErrorCategory]::InvalidData
                                ErrorId = 'PathNotValidForProvider'
                                TargetObject = $_
                                RecommendedAction = "Use a provider qualified path for the specified provider '$($providerInfo.Name)'."
                            }
                            throw (New-ADTErrorRecord @naerParams)
                        }
                    }

                    $qualifiedPath = "$($providerInfo.ModuleName)\$($providerInfo.Name)::$providerPath"

                    # Try get the item by specifying the the qualified path and invoking the Item.Get method.
                    $(
                        try
                        {
                            # Some providers do not throw when the path does not exist, so we do if no items are returned.
                            if (!([string[]]$items = $ExecutionContext.SessionState.InvokeProvider.Item.Get($qualifiedPath, $Force, $isLiteralPath) | & { process { $_.PSPath } }))
                            {
                                throw [System.Management.Automation.ItemNotFoundException]::new("The given path `"$_`" does not resolve to any item or is not a valid path.")
                            }
                            return $items
                        }
                        catch [System.Management.Automation.ItemNotFoundException]
                        {
                            # Ignore issues with non-existent paths if the parameter is set and the path is not a wildcard or literal path.
                            if ($IncludeNonExistent -and ($isLiteralPath -or ![WildcardPattern]::ContainsWildcardCharacters($qualifiedPath))) {
                                return $qualifiedPath
                            }
                            # If the path contains wildcard resolution it should not throw an error.
                            elseif (!$IncludeNonExistent -and ($isLiteralPath -or ![WildcardPattern]::ContainsWildcardCharacters($qualifiedPath))) {
                                $naerParams = @{
                                    Exception = $_
                                    Category = [System.Management.Automation.ErrorCategory]::ObjectNotFound
                                    ErrorId = 'PathNotFound'
                                    TargetObject = $qualifiedPath
                                    RecommendedAction = "Check if the path exists or if it is a valid path."
                                }
                                throw (New-ADTErrorRecord @naerParams)
                            }
                        }
                        catch
                        {
                            $naerParams = @{
                                Exception = $_
                                Category = [System.Management.Automation.ErrorCategory]::InvalidOperation
                                ErrorId = 'PathResolutionError'
                                TargetObject = $qualifiedPath
                                RecommendedAction = "Check if the path is valid or if the provider supports the operation."
                            }
                            throw (New-ADTErrorRecord @naerParams)
                        }
                    ) |
                    & {
                        # Filter the items based on the path type, filter, exclude, and include parameters.
                        process
                        {
                            # Filter out items that do not match the specified path type.
                            if ($PathType -ne [Microsoft.PowerShell.Commands.TestPathType]::Any -and
                                $ExecutionContext.SessionState.Item.IsContainer($_) -ne ($PathType -eq [Microsoft.PowerShell.Commands.TestPathType]::Container)
                            ) { return }

                            # Filter out items that do not match the specified filter, exclude, or include parameters.
                            $pathLeaf = $ExecutionContext.SessionState.Path.ParseChildName($_)
                            if (($Filter -and ($pathLeaf -notlike $Filter)) -or
                                ($Exclude -and $true -in ($Exclude | & { process { $pathLeaf -like $Filter } })) -or
                                ($Include -and $true -notin ($Include | & { process { $pathLeaf -like $Filter } }))
                            ) { return }

                            # Return as the desired path format.
                            if ($AsProviderPath) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_) } else { return $_ }
                        }
                    }
                }
            }
        }
        catch
        {
            Invoke-ADTFunctionErrorHandler -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -ErrorRecord $_
        }
    }

    end
    {
        Complete-ADTFunction -Cmdlet $PSCmdlet
    }
}
