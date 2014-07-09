del c:\w3wpProcs.txt

function Get-Executables
{
    [CmdletBinding()]
    param (
        [string] $ProcessName = $(throw "Please specify a process name.")
    )

    process
    {
        # Retrieve instance of Win32_Process based on the process name passed into the function
        $Procs = Get-WmiObject -Namespace root\cimv2 -Query "select * from Win32_Process where Name = '$ProcessName'"
		$Pc = $Procs.Count;
		
        # If there are no processes returned from the query, then simply exit the function
        if (-not $Procs)
        {
            Write-Host -Object "No processes were found named $ProcessName";
            return;
        }
        # If one process is found, get the value of __PATH, which we will use for our next query
        elseif (@($Procs).Count -eq 1)
        {
            Write-Verbose "One process was found named $ProcessName";
			Add-Content c:\w3wpProcs.txt "";
			Add-Content c:\w3wpProcs.txt "============================================================================";
			Add-Content c:\w3wpProcs.txt "One process was found named $ProcessName";
            $ProcPath = @($Procs)[0].__PATH;
            Write-Verbose "Proc path is $ProcPath";
			Add-Content c:\w3wpProcs.txt "Proc path is $ProcPath";
			Add-Content c:\w3wpProcs.txt "============================================================================";
			Add-Content c:\w3wpProcs.txt "";
			Process-Query;
        }
        # If there is more than one process, use the process at index 0, for the time being
        elseif ($Procs.Count -gt 1)
        {
			$i = 0
			while ($i -le $Pc - 1)
			{
				Write-Host -Object "More than one process was found named $ProcessName";
				Add-Content c:\w3wpProcs.txt "";
				Add-Content c:\w3wpProcs.txt "============================================================================";
				Add-Content c:\w3wpProcs.txt "More than one process was found named $ProcessName";
				$ProcPath = @($Procs)[$i].__PATH;
				Write-Host -Object "Using process with path: $ProcPath";
				Add-Content c:\w3wpProcs.txt "Using process with path: $ProcPath";
				Add-Content c:\w3wpProcs.txt "============================================================================";
				Add-Content c:\w3wpProcs.txt "";
				Process-Query;
				$i++;
			}
        }
		#Process-Query;
	}

    # Do a little clean-up work
    end
    {
        Write-Verbose "End: Cleaning up variables used for function"
        Remove-Item -ErrorAction SilentlyContinue -Path variable:ExeFile,variable:ProcessName,variable:ProcExe,
        variable:ProcExes,variable:ProcPath,variable:ProcQuery,variable:Procs
    }
}

Clear-Host;

function Process-Query
{
	[CmdletBinding()]
	
	Param()
	
	process
	{
		# Get the CIM_ProcessExecutable instances for the process we retrieved
        $ProcPath;
        $ProcQuery = "select * from CIM_ProcessExecutable where Dependent = '$ProcPath'".Replace('\','\\');

        Write-Verbose $ProcQuery
        $ProcExes = Get-WmiObject -Namespace root\cimv2 -Query $ProcQuery;

        # If there are instances of CIM_ProcessExecutable for the specified process, go ahead and grab the important properties
        if ($ProcExes)
        {
            foreach ($ProcExe in $ProcExes)
            {
                # Use the [wmi] type accelerator to retrieve an instance of CIM_DataFile from the WMI __PATH in the Antecentdent property
                $ExeFile = [wmi]"$($ProcExe.Antecedent)"
                # If the WMI instance we just retrieve "IS A" (think WMI operator) CIM_DataFile, then write properties to console
                if ($ExeFile.__CLASS -eq 'CIM_DataFile')
                {
                    Select-Object -InputObject $ExeFile -Property FileName,Extension,Manufacturer,Version -OutVariable $Executables | Add-Content c:\w3wpProcs.txt;
                }
            }
        }
	}
}
# Call the function we just defined, with its single parameter
Get-Executables -ProcessName w3wp.exe;