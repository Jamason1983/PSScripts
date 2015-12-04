function Get-IIS6FTPVirtualDirectories
{
[CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$false,
          HelpMessage='Please enter a valid fully qualafied path to an IIS 6 metabase XML file. Usualy found at "c:\WINDOWS\system32\inetsrv" on the server that is hosting IIS.')]
        [System.IO.FileInfo]$Path
    )
    Begin
    {
        #Set variables to be used for function
        $FTPVDList = @()
        $FTPSiteList = @()

        #Get Content from metabase xml
        [xml]$XML = Get-Content -Path $Path

        #extract data from XML to variable for use.
        $FTPVirtualDirectories = $XML.configuration.MBProperty.IIsFtpVirtualDir
        $FTPSites = $XML.configuration.MBProperty.IIsFtpServer

        #Get the Site ids and names.
        foreach($FTPSite in $FTPSites)
        {
            $SiteID = $FTPSite.Location.SubString(13)
            $SiteName = $FTPSite.ServerComment

            $Object = New-Object PSObject                                       
            $Object | add-member Noteproperty -Name SiteId -Value $SiteID
            $Object | add-member Noteproperty -Name Name -Value $SiteName           
            $FTPSiteList += $Object
        }
    }
    Process
    {
        foreach($VirtualDirectorie in $FTPVirtualDirectories)
        {
            $FTPSiteId = $VirtualDirectorie.Location.SubString(0,$VirtualDirectorie.Location.IndexOf("/root")).SubString(13)
            $FTPVDName = $VirtualDirectorie.Location.SubString($VirtualDirectorie.Location.LastIndexOf("/") + 1)
            $FTPPath = $VirtualDirectorie.Path

            #Find Site name from site id
            $Site = $FTPSiteList|?{$_.SiteId -eq $FTPSiteId}


            $Object = New-Object PSObject                                       
            $Object | add-member Noteproperty -Name SiteId -Value $FTPSiteId
            $object | add-member Noteproperty -Name SiteName -Value $Site.Name
            $Object | add-member Noteproperty -Name VDName -Value $FTPVDName
            $Object | add-member Noteproperty -Name VDPath -Value $FTPPath
            $FTPVDList += $Object
        }
    }
    End
    {
        Write-Output $FTPVDList
    }
}

function Add-VirtualDirectoriestoFTP
{
[CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [psobject[]]$VDList
    )
    Begin
    {
        $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
        $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
        $IsAdmin=$prp.IsInRole($adm)
        if ($IsAdmin)
        {
            #import Module
            Import-Module WebAdministration -ErrorAction Stop
        }
        else
        {
            Write-Error -Category PermissionDenied -Message "You need to run this function in an elevated powershell prompt. Please open a new powershell window as an admin." -ErrorAction Stop
        }
        
        #get list of FTP Sites.
        $Sites = Get-Website
    }
    Process
    {
        foreach($Site in $Sites)
        {
            $SiteVDstoAdd = $VDList|Where-Object SiteName -eq $Site.Name|Where-Object VDName -NE "root"
            foreach($SiteVD in $SiteVDstoAdd)
            {
                #See if VD is already added
                $GetSiteVD = Get-WebVirtualDirectory -Site $Site.Name -Name $SiteVD.VDName -ErrorAction SilentlyContinue
                if(!$GetSiteVD)
                {
                    try
                    {
                        New-WebVirtualDirectory -Site $Site.Name -Name $SiteVD.VDName -PhysicalPath $SiteVD.VDPath
                        $VDCreateStatus = "Created"
                    }
                    catch
                    {
                        if($Error[0].Exception -match "Parameter 'PhysicalPath' should point to existing path.")
                        {
                            #Write-Warning -Message "FTP VD [PhysicalPath NOT FOUND] Site: $($Site.Name) Name: $($SiteVD.VDName) PhysicalPath: $($SiteVD.VDPath)"
                            $VDCreateStatus = "Fail: Path not found"
                        }
                        else
                        {
                            Write-Error -Exception $Error[0]
                            $VDCreateStatus = "Fail: $($Error[0].Exception)"
                        }
                    }
                }
                else
                {
                    #Write-Warning "FTP VD [ALREADY EXIST] Site: $($Site.Name) Name: $($SiteVD.VDName) PhysicalPath: $($SiteVD.VDPath)"
                    $VDCreateStatus = "VD already exist"
                }

                #Write object to pipeline
                $Object = New-Object PSObject
                $object | add-member Noteproperty -Name SiteName -Value $Site.Name
                $Object | add-member Noteproperty -Name VDName -Value $SiteVD.VDName
                $Object | add-member Noteproperty -Name VDPath -Value $SiteVD.VDPath
                $Object | add-member Noteproperty -Name Status -Value $VDCreateStatus
                Write-Output $Object
            }
        }
    }
}

Get-IIS6FTPVirtualDirectories -Path "\\server.domain.com\c$\WINDOWS\system32\inetsrv\MetaBase.xml"|Add-VirtualDirectoriestoFTP
