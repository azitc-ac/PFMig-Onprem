<#
    Functions for Public Folder migration (on-prem, WIA)
    Authors (original/adapted): Glen Scales, Alexander Zarenko
    Refactor: 2026-01
    Notes:
      - PowerShell 5.1 compatible
      - Uses EWS Managed API (WIA/Default Credentials)
      - TLS validation bypass is active (same as original)
#>

function Load-EWSManagedAPI {
    [CmdletBinding()]
    param()

    # Check if EWS assembly is already loaded in current AppDomain (don't reload)
    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() |
              Where-Object { $_.GetName().Name -eq 'Microsoft.Exchange.WebServices' }
    if ($loaded) {
        Write-Verbose ("EWS already loaded: {0}" -f ($loaded | Select-Object -First 1 -Expand Location))
        return
    }

    # Otherwise: load highest version from Registry (as before)
    $key = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'
    $ewsKey = Get-ChildItem -Path $key -ErrorAction SilentlyContinue |
              Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name
    if ($ewsKey) {
        $installDir = (Get-ItemProperty -Path "Registry::$ewsKey" -ErrorAction SilentlyContinue).'Install Directory'
        $dll = Join-Path $installDir 'Microsoft.Exchange.WebServices.dll'
        if (Test-Path $dll) {
            Import-Module $dll -ErrorAction Stop
            return
        }
    }
    Write-Error "EWS Managed API (>=1.2) not found. Please install."
    throw
}

function Handle-SSL {
    [CmdletBinding()]
    param()
    # Keep original: trust-all TLS (for lab/compat) – security risk in prod
    $Provider = New-Object Microsoft.CSharp.CSharpCodeProvider
    $Params = New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable = $False
    $Params.GenerateInMemory   = $True
    $Params.IncludeDebugInformation = $False
    [void]$Params.ReferencedAssemblies.Add("System.DLL")
$TASource=@'
namespace Local.ToolkitExtensions.Net.CertificatePolicy{
  public class TrustAll : System.Net.ICertificatePolicy {
    public TrustAll() { }
    public bool CheckValidationResult(System.Net.ServicePoint sp,
      System.Security.Cryptography.X509Certificates.X509Certificate cert,
      System.Net.WebRequest req, int problem) {
      return true;
    }
  }
}
'@
    $TAResults   = $Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly  = $TAResults.CompiledAssembly
    $TrustAll    = $TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    [System.Net.ServicePointManager]::CertificatePolicy = $TrustAll
}

function Connect-Exchange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$MailboxName,
        [Parameter()][string]$Url
    )
    Load-EWSManagedAPI
    Handle-SSL

    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
    $service.UseDefaultCredentials = $true

    if ($Url) {
        $uri = [System.Uri]$Url
        $service.Url = $uri
    } else {
        # 1) Try Autodiscover with WIA
        try {
            $service.AutodiscoverUrl($MailboxName, { $true })
        } catch {
            # 2) Fallback: Script runs on Exchange server itself.
            #    Autodiscover often fails (no SCP, loopback, no DNS record).
            #    Use local EWS endpoint directly. Try localhost first
            #    because it avoids Kerberos loopback 401.
            Write-Warning ("Autodiscover failed ({0}). Trying local EWS endpoints." -f $_.Exception.Message)

            $fqdn = $null
            try { $fqdn = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName } catch {}

            $candidates = New-Object System.Collections.Generic.List[string]
            $candidates.Add("https://localhost/EWS/Exchange.asmx")
            if ($fqdn) { $candidates.Add(("https://{0}/EWS/Exchange.asmx" -f $fqdn)) }

            $connected = $false
            foreach ($candidate in $candidates) {
                try {
                    $service.Url = [System.Uri]$candidate
                    # Inexpensive validation: bind root folder of admin mailbox
                    $testId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
                        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName
                    )
                    [void][Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $testId)
                    Write-Host ("Connected to EWS via fallback URL [{0}]" -f $candidate) -ForegroundColor DarkGray
                    $connected = $true
                    break
                } catch {
                    Write-Verbose ("Fallback URL [{0}] failed: {1}" -f $candidate, $_.Exception.Message)
                    $service.Url = $null
                }
            }
            if (-not $connected) {
                throw ("Connection to EWS failed. Please specify -Url explicitly (e.g., https://localhost/EWS/Exchange.asmx). Mailbox: '{0}'" -f $MailboxName)
            }
        }
    }

    if (-not $service.Url) { throw ("Error connecting to EWS (Mailbox '{0}')." -f $MailboxName) }
    return $service
}

function Get-TargetOU {
    <#
      Determines target OU (complete DN) for mailbox creation.
      Logic:
        - If config file exists with saved DN, ask whether to use it for this run (Y/N)
        - If "No" or config missing, use Out-GridView selection
        - Selected OU is saved to config file
      Returns: DN string or $null on cancel
    #>
    [CmdletBinding()]
    param(
        [Parameter()][string]$ConfigPath,
        [Parameter()][string]$Title = "Mailbox creation: Select OU for disabled user account"
    )

    # 1) Read saved DN (if config exists)
    $savedDn = $null
    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        try {
            $raw = (Get-Content -Path $ConfigPath -Raw -ErrorAction Stop).Trim()
            if (-not [string]::IsNullOrWhiteSpace($raw)) { $savedDn = $raw }
        } catch {
            Write-Warning ("Could not read OU config '{0}': {1}" -f $ConfigPath, $_.Exception.Message)
        }
    }

    # 2) Offer saved DN
    if ($savedDn) {
        Write-Host ("Saved target OU found:`n    {0}" -f $savedDn) -ForegroundColor Cyan
        $answer = Read-Host "Use this OU for current run? [Y/N]"
        if ($answer -match '^(?i)\s*(j|ja|y|yes)\s*$') {
            return $savedDn
        }
        Write-Host "OK - selecting new OU via Out-GridView..." -ForegroundColor Yellow
    }

    # 3) OGV-Auswahl
    $ou = Get-ADOrganizationalUnit -Filter * | Select-Object DistinguishedName, Name |
          Out-GridView -PassThru -Title $Title
    if (-not $ou) { return $null }
    if ($ou -is [System.Array]) { $ou = $ou | Select-Object -First 1 }
    $ouDn = $ou.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($ouDn)) { return $null }

    # 4) Persist selection
    if ($ConfigPath) {
        try {
            Set-Content -Path $ConfigPath -Value $ouDn -Encoding UTF8 -ErrorAction Stop
            Write-Host ("Target OU saved in [{0}]" -f $ConfigPath) -ForegroundColor DarkGray
        } catch {
            Write-Warning ("Could not save OU to config '{0}': {1}" -f $ConfigPath, $_.Exception.Message)
        }
    }
    return $ouDn
}

function Test-EwsMailboxAccess {
    <#
      Actively checks whether the given EWS service can access the mailbox
      (bind to MsgFolderRoot). With retry because:
        - Newly created mailboxes must be provisioned in the Store first
        - RBAC/Impersonation configuration takes a few minutes to take effect
      Distinguishes TRANSIENT errors (wait and retry) from
      PERMANENT auth/impersonation errors (abort with diagnostic message immediately).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory=$true)][string]$MailboxSmtp,
        [Parameter()][int]$MaxAttempts = 12,
        [Parameter()][int]$DelaySeconds = 10,
        [Parameter()][string]$Context = 'Target mailbox'
    )

    $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxSmtp
    )

    $lastMsg = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            [void][Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $rootId)
            Write-Verbose ("EWS access to {0} [{1}] confirmed (attempt {2})." -f $Context, $MailboxSmtp, $attempt)
            return $true
        } catch {
            $msg = $_.Exception.Message
            $lastMsg = $msg
            # Detect permanent auth/impersonation errors -> don't retry indefinitely
            $isAuth = $msg -match '(?i)401|Unauthorized|nicht autorisiert|ImpersonateUserDenied|ImpersonationFailed|ImpersonateUser|AccessDenied|Zugriff verweigert'
            if ($isAuth) {
                throw ("EWS access to {0} [{1}] denied (authentication/impersonation). " -f $Context, $MailboxSmtp) +
                      "Please verify: Is ApplicationImpersonation role assigned to the running account? " +
                      "(New-ManagementRoleAssignment -Role ApplicationImpersonation -User <admin>) " +
                      "RBAC cache may take ~15 minutes to propagate. Original error: $msg"
            }
            # Transient (mailbox not yet provisioned / store temporarily unavailable).
            # Discreet wait message on first attempt, details only with -Verbose.
            if ($attempt -eq 1) {
                Write-Host ("  Waiting for {0} [{1}] provisioning (up to {2}s)..." -f $Context, $MailboxSmtp, ($MaxAttempts * $DelaySeconds)) -ForegroundColor DarkGray
            }
            Write-Verbose ("EWS access not yet available (attempt {0}/{1}): {2}" -f $attempt, $MaxAttempts, $msg)
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    $secs = $MaxAttempts * $DelaySeconds
    if ($lastMsg -match '(?i)temporarily unavailable|MailboxStoreUnavailable|failed to get the correct properties|voruebergehend nicht verfuegbar') {
        # Store/database reports unavailability -> not a permission/impersonation problem.
        throw ("EWS access to {0} [{1}] failed after {2}s: mailbox store/database is 'temporarily unavailable'. " -f $Context, $MailboxSmtp, $secs) +
              "This is NOT a permission/impersonation problem. Please verify: " +
              "Database mounted? (Get-MailboxDatabase -Status | ft Name,Mounted)  " +
              "Service 'Microsoft Exchange Information Store' running? (Get-Service MSExchangeIS)  " +
              "Mailbox visible in store? (Get-MailboxStatistics -Identity '$MailboxSmtp')  " +
              "The mailbox now exists — re-running skips creation and tries access directly. " +
              "Note: If Exchange runs on a Domain Controller, store timeouts are common (unsupported setup). " +
              "Last error: $lastMsg"
    }
    throw ("EWS access to {0} [{1}] failed after {2}s. " -f $Context, $MailboxSmtp, $secs) +
          "Mailbox may not be fully provisioned yet, or permission/impersonation is missing. Last error: $lastMsg"
}

function Resolve-AutodiscoverUrl {
    <#
      Resolves the Autodiscover .svc URL. Order of precedence:
        1) Explicitly passed value
        2) From the (working) EWS service host -> same CAS serves
           EWS and Autodiscover, so reliable once EWS is connected
        3) If localhost/127.0.0.1: use FQDN of machine instead (avoids
           Kerberos loopback 401 on Autodiscover SOAP call)
      Returns: URL string or $null
    #>
    [CmdletBinding()]
    param(
        [Parameter()][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter()][string]$AutodiscoverUrl
    )
    if ($AutodiscoverUrl) { return $AutodiscoverUrl }

    $hostName = $null
    if ($Service -and $Service.Url) { $hostName = $Service.Url.Host }
    if (-not $hostName -or $hostName -eq 'localhost' -or $hostName -eq '127.0.0.1') {
        try { $hostName = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName } catch {}
    }
    if ($hostName) { return ("https://{0}/autodiscover/autodiscover.svc" -f $hostName) }
    return $null
}

function Get-PublicFolderRoutingHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory=$true)][string]$MailboxName,
        [Parameter(Mandatory=$true)][string]$Header,
        [Parameter()][string]$AutodiscoverUrl
    )
    if ($Header -ne 'X-AnchorMailbox') { return }
    # Header schon gesetzt? Nichts zu tun.
    if ($Service.HttpHeaders.$Header) { return }

    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $resolvedUrl = Resolve-AutodiscoverUrl -Service $Service -AutodiscoverUrl $AutodiscoverUrl

    # Try multiple strategies in sequence (self-healing if one fails):
    #   1) Explicit/derived Autodiscover URL (no SCP)
    #   2) SCP lookup in AD (very reliable on domain-joined Exchange server)
    $strategies = New-Object System.Collections.ArrayList
    if ($resolvedUrl) { [void]$strategies.Add(@{ Url = $resolvedUrl; Scp = $false }) }
    [void]$strategies.Add(@{ Url = $null; Scp = $true })

    $pfi = $null
    $lastErr = $null
    foreach ($s in $strategies) {
        try {
            $ads = New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion)
            $ads.UseDefaultCredentials = $true
            $ads.EnableScpLookup = [bool]$s.Scp
            $ads.RedirectionUrlValidationCallback = { $true }
            $ads.PreAuthenticate = $true
            $ads.KeepAlive = $false
            if ($s.Url) { $ads.Url = [System.Uri]$s.Url }

            $gsp = $ads.GetUserSettings($MailboxName, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation)
            $val = $null
            if ($gsp.Settings.TryGetValue([Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::PublicFolderInformation, [ref]$val)) {
                $pfi = $val
                Write-Verbose ("PublicFolderInformation via Autodiscover (Url={0}, Scp={1}): {2}" -f $s.Url, $s.Scp, $pfi)
                break
            }
        } catch {
            $lastErr = $_.Exception.Message
            Write-Verbose ("Autodiscover attempt (Url={0}, Scp={1}) failed: {2}" -f $s.Url, $s.Scp, $lastErr)
        }
    }

    if ($pfi) {
        if (-not $Service.HttpHeaders.$Header) { $Service.HttpHeaders.Add($Header, $pfi) }
    } else {
        throw ("Could not determine PublicFolderInformation (X-AnchorMailbox) via Autodiscover. " +
               "Please specify -AutodiscoverUrl explicitly (e.g., https://ex.test.zarenko.net/autodiscover/autodiscover.svc). " +
               "Last error: $lastErr")
    }
}

function Get-PublicFolderContentRoutingHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

        [Parameter(Mandatory=$true)]
        [string]$MailboxName,

        [Parameter(Mandatory=$true)]
        [string]$PfAddress,

        [Parameter()]
        [string]$AutodiscoverUrl
    )

    process {
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
        $ads = New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($ExchangeVersion)
        $ads.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        $ads.UseDefaultCredentials = $true
        $ads.EnableScpLookup = $false
        $ads.RedirectionUrlValidationCallback = { $true }
        $ads.PreAuthenticate = $true
        $ads.KeepAlive = $false

        # Prefer deriving URL from EWS host (reliable once EWS is connected)
        $resolvedUrl = Resolve-AutodiscoverUrl -Service $Service -AutodiscoverUrl $AutodiscoverUrl
        if ($resolvedUrl) {
            $ads.Url = [System.Uri]$resolvedUrl
        } else {
            # Last resort: let Autodiscover discover itself (SCP)
            $ads.EnableScpLookup = $true
            $null = $ads.GetUserSettings($MailboxName, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress)
        }

        if (-not $ads.Url) {
            throw "Autodiscover URL could not be determined. Provide -AutodiscoverUrl (e.g., https://ex.test.zarenko.net/autodiscover/autodiscover.svc)."
        }

        $xml = '<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006"><Request>' +
               ("<EMailAddress>{0}</EMailAddress>" -f $PfAddress) +
               '<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>' +
               '</Request></Autodiscover>'

        $req = [System.Net.HttpWebRequest]::Create($ads.Url.ToString().Replace(".svc",".xml"))
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($xml)
        $req.ContentLength = $bytes.Length
        $req.ContentType   = "text/xml"
        $req.UserAgent     = "Microsoft Office/16.0 (Windows NT 6.3; Microsoft Outlook 16.0.6001; Pro)"
        $req.Headers.Add("Translate","F")
        $req.Method        = "POST"
        $req.Credentials   = [System.Net.CredentialCache]::DefaultNetworkCredentials

        $stream = $req.GetRequestStream()
        $stream.Write($bytes,0,$bytes.Length)
        $stream.Close()

        $req.AllowAutoRedirect = $true   # Bugfix (dein Original hatte $truee)

        $resp = $req.GetResponse().GetResponseStream()
        $sr   = New-Object System.IO.StreamReader($resp)
        [xml]$xmlResp = $sr.ReadToEnd()

        if ($xmlResp.Autodiscover.Response.User.AutoDiscoverSMTPAddress) {
            $anchor = $xmlResp.Autodiscover.Response.User.AutoDiscoverSMTPAddress
            Write-Verbose ("Public Folder Content Routing Information Header : {0}" -f $anchor)
            $Service.HttpHeaders["X-AnchorMailbox"]      = $anchor
            $Service.HttpHeaders["X-PublicFolderMailbox"] = $anchor
        }
    }
}

function Get-FolderFromPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$FolderPath,
        [Parameter(Mandatory=$true)][string]$MailboxName,
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter()][Microsoft.Exchange.WebServices.Data.PropertySet]$PropertySet
    )

    # Bind to MsgFolderRoot of target mailbox
    $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName
    )

    # EARLY RETURN for root
    if ([string]::IsNullOrEmpty($FolderPath) -or $FolderPath -eq '\') {
        return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$rootId)
    }

    $tf = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$rootId)

    # Split path safely: remove leading/trailing "\" and empty segments.
    # Explicit [string[]] typing + cast per element: prevents PS 5.1
    # from coercing single-element Where-Object output to scalar/boolean.
    [string[]]$segments = $FolderPath.Trim('\').Split('\') | Where-Object { $_ -ne '' } | ForEach-Object { [string]$_ }

    foreach ($seg in $segments) {
        $fv = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        if ($PropertySet) { $fv.PropertySet = $PropertySet }
        $sf = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $seg
        )
        # Retry loop: EWS doesn't always see newly created folders immediately
        $res = $null
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            $res = $Service.FindFolders($tf.Id, $sf, $fv)
            if ($res.TotalCount -gt 0) { break }
            Write-Verbose ("Get-FolderFromPath: Segment '{0}' not yet visible, waiting 2s (attempt {1}/5)..." -f $seg, $attempt)
            Start-Sleep -Seconds 2
        }
        if ($res.TotalCount -gt 0) {
            foreach ($f in $res.Folders) { $tf = $f }
        } else {
            Write-Warning ("Get-FolderFromPath: Segment '{0}' in path '{1}' not found." -f $seg, $FolderPath)
            return $null
        }
    }

    if ($tf -ne $null) { return [Microsoft.Exchange.WebServices.Data.Folder]$tf }
    throw "Folder Not found"
}

function PublicFolderIdFromPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

        [Parameter(Mandatory=$true)]
        [string]$FolderPath,

        [Parameter(Mandatory=$true)]
        [string]$SmtpAddress,

        [Parameter()]
        [string]$AutodiscoverUrl
    )

    process {
        # 1) Ensure X-AnchorMailbox header is set
        if (-not $Service.HttpHeaders['X-AnchorMailbox']) {
            Get-PublicFolderRoutingHeader -Service $Service -MailboxName $SmtpAddress -Header 'X-AnchorMailbox' -AutodiscoverUrl $AutodiscoverUrl
        }

        # 2) Bind PublicFoldersRoot and prepare PropertySet
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot
        )

        $ps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
        )

        # Optionally add PR_REPLICA_LIST (may fail with mixed EWS DLL versions)
        $PR_REPLICA_LIST = $null
        try {
            $PR_REPLICA_LIST = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
                0x6698, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary
            )
            # Explicit cast – works fine with homogeneous assembly, caught on mismatch
            [void]$ps.Add([Microsoft.Exchange.WebServices.Data.PropertyDefinitionBase]$PR_REPLICA_LIST)
        } catch {
            Write-Verbose "Could not add PR_REPLICA_LIST to PropertySet (possible EWS type conflict). Continuing without it."
            $PR_REPLICA_LIST = $null
        }

        $tf = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderId, $ps)

        # 3) Root case: return immediately
        if ([string]::IsNullOrEmpty($FolderPath) -or $FolderPath -eq '\') {
            # Optional: derive routing header from REPLICA_LIST (if available)
            if ($PR_REPLICA_LIST) {
                $val = $null
                if ($tf.TryGetProperty($PR_REPLICA_LIST, [ref]$val)) {
                    $guid = [System.Text.Encoding]::ASCII.GetString($val, 0, 36)
                    $addr = New-Object System.Net.Mail.MailAddress($Service.HttpHeaders['X-AnchorMailbox'])
                    $pfHeader = $guid + '@' + $addr.Host
                    Write-Verbose ("Root Public Folder Routing Information Header : {0}" -f $pfHeader)
                    if (-not $Service.HttpHeaders.'X-PublicFolderMailbox') {
                        $Service.HttpHeaders.Add('X-PublicFolderMailbox', $pfHeader)
                    }
                } else {
                    Write-Verbose "PR_REPLICA_LIST not available at root – using only X-AnchorMailbox."
                }
            }
            return $tf.Id.UniqueId.ToString()
        }

        # 4) Search subtree – split path safely (see Get-FolderFromPath)
        [string[]]$segments = $FolderPath.Trim('\').Split('\') | Where-Object { $_ -ne '' } | ForEach-Object { [string]$_ }

        foreach ($seg in $segments) {
            $fv = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
            $fv.PropertySet = $ps
            $sf = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
                [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $seg
            )
            $res = $Service.FindFolders($tf.Id, $sf, $fv)

            if ($res.TotalCount -gt 0) {
                foreach ($f in $res.Folders) { $tf = $f }
            } else {
                Write-Error "Error Folder Not Found (Segment: '$seg' in path '$FolderPath')"
                return $null
            }
        }

        # 5) After successful resolution: optionally add content routing header
        if ($PR_REPLICA_LIST) {
            $val = $null
            if ($tf.TryGetProperty($PR_REPLICA_LIST, [ref]$val)) {
                $guid = [System.Text.Encoding]::ASCII.GetString($val, 0, 36)
                $addr = New-Object System.Net.Mail.MailAddress($Service.HttpHeaders['X-AnchorMailbox'])
                $pfHeader = $guid + '@' + $addr.Host
                Write-Verbose ("Target Public Folder Routing Information Header : {0}" -f $pfHeader)

                # Set X-AnchorMailbox/X-PublicFolderMailbox for content routing (Autodiscover XML)
                Get-PublicFolderContentRoutingHeader -Service $Service -MailboxName $SmtpAddress -PfAddress $pfHeader -AutodiscoverUrl $AutodiscoverUrl
            } else {
                Write-Verbose "PR_REPLICA_LIST not available at target folder – using only X-AnchorMailbox."
            }
        } else {
            Write-Verbose "PR_REPLICA_LIST not set – using only X-AnchorMailbox."
        }

        # 6) Return UniqueId as string
        return $tf.Id.UniqueId.ToString()
    }
}

function Create-Folder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$MailboxName,
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory=$true)][string]$NewFolderName,
        [Parameter()][string]$ParentFolder,
        [Parameter()][string]$FolderClass
    )

    $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
    $newFolder.DisplayName = $NewFolderName
    $newFolder.FolderClass = if ([string]::IsNullOrEmpty($FolderClass)) { 'IPF.Note' } else { $FolderClass }

    # Bind root directly if parent is empty or "\"
    if ([string]::IsNullOrEmpty($ParentFolder) -or $ParentFolder -eq '\') {
        $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName
        )
        $EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$rootId)
    } else {
        # Resolve parent path cleanly
        $EWSParentFolder = Get-FolderFromPath -MailboxName $MailboxName -Service $Service -FolderPath $ParentFolder
        if (-not $EWSParentFolder) { throw "Parent folder '$ParentFolder' not found in target mailbox '$MailboxName'." }
    }

    # Check existence
    $fv = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $sf = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
        [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $NewFolderName
    )
    $res = $Service.FindFolders($EWSParentFolder.Id, $sf, $fv)

    if ($res.TotalCount -eq 0) {
        $newFolder.Save($EWSParentFolder.Id)
        Write-Verbose ("Folder [{0}] Created (Parent [{1}])" -f $NewFolderName, $(if ($ParentFolder) { $ParentFolder } else { '\' }))
        $created = $true
    } else {
        Write-Verbose ("Folder [{0}] already exists (Parent [{1}])." -f $NewFolderName, $(if ($ParentFolder) { $ParentFolder } else { '\' }))
        $created = $false
    }

    # Return status so caller can log what was newly created
    return [pscustomobject]@{
        FolderName = $NewFolderName
        Parent     = $(if ($ParentFolder) { $ParentFolder } else { '\' })
        Created    = $created
    }
}

function Get-PublicFolderItems {
    <#
      Refactored version: Copies items from PF path to target mailbox folder.
      - No global variables required
      - Supports ShouldProcess
      - CSV de-dup uses Item.Id.UniqueId

      IMPORTANT: $Service is used exclusively for Public Folder EWS access
                 (receives X-AnchorMailbox / X-PublicFolderMailbox routing headers).
                 $TargetService is a separate, clean EWS connection for
                 accessing the target mailbox (without PF routing headers).
                 Both instances MUST be created separately (Connect-Exchange).
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        # EWS service for Public Folder access (receives PF routing headers)
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        # EWS service for target mailbox access (WITHOUT PF routing headers!)
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$TargetService,
        [Parameter(Mandatory=$true)][string]$AdminMailboxSmtp,
        [Parameter(Mandatory=$true)][string]$PublicFolderPath,
        [Parameter(Mandatory=$true)][string]$TargetMailboxSmtp,
        [Parameter(Mandatory=$true)][string]$ParentPath,
        [Parameter(Mandatory=$true)][string]$FolderName,
        [Parameter()][string]$FolderClass,
        [Parameter()][string]$AutodiscoverUrl,
        [Parameter()][switch]$DoNotCopyItems
    )
    # Pre-calculate path in target mailbox (for reporting even on errors)
    if ([string]::IsNullOrEmpty($ParentPath) -or $ParentPath -eq '\') {
        $targetFolderPath = '\' + $FolderName
    } else {
        $targetFolderPath = $ParentPath.TrimEnd('\') + '\' + $FolderName
    }

    # Result object returned at the end (even on errors)
    # so orchestrator can generate meaningful final accounting
    $result = [pscustomobject]@{
        SourcePath    = $PublicFolderPath
        TargetPath    = $targetFolderPath
        FolderCreated = $false
        ItemsCopied   = 0
        ItemsSkipped  = 0
        ItemsFailed   = 0
        Status        = 'OK'
        Error         = $null
    }

    # --- SOURCE: Determine public folder ID (PF service) ---
    try {
        Get-PublicFolderRoutingHeader -Service $Service -MailboxName $AdminMailboxSmtp -Header "X-AnchorMailbox" -AutodiscoverUrl $AutodiscoverUrl
        $fldId = PublicFolderIdFromPath -Service $Service -FolderPath $PublicFolderPath -SmtpAddress $AdminMailboxSmtp -AutodiscoverUrl $AutodiscoverUrl
        if (-not $fldId) { throw "PublicFolderIdFromPath returned no ID." }
        $subFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($fldId)
    } catch {
        $m = $_.Exception.Message
        $hint = if ($m -match '(?i)401|Unauthorized|nicht autorisiert|AccessDenied|Zugriff verweigert') {
            " -> SOURCE access denied: does the running account have read permissions on the public folder?"
        } else { "" }
        $result.Status = 'SourceFailed'
        $result.Error  = $m + $hint
        Write-Error ("Source PF '{0}' not accessible: {1}{2}" -f $PublicFolderPath, $m, $hint)
        return $result
    }

    # --- TARGET: Create folder (TargetService, NO PF routing) ---
    try {
        $cf = Create-Folder -MailboxName $TargetMailboxSmtp -Service $TargetService -NewFolderName $FolderName -ParentFolder $ParentPath -FolderClass $FolderClass
        if ($cf) { $result.FolderCreated = [bool]$cf.Created }
    } catch {
        $m = $_.Exception.Message
        $hint = if ($m -match '(?i)401|Unauthorized|nicht autorisiert|Impersonat|AccessDenied|Zugriff verweigert') {
            " -> TARGET access denied: is ApplicationImpersonation role set? Is mailbox provisioned?"
        } else { "" }
        $result.Status = 'TargetFolderFailed'
        $result.Error  = $m + $hint
        Write-Error ("Target folder '{0}' could not be created: {1}{2}" -f $targetFolderPath, $m, $hint)
        return $result
    }

    # Resolve target folder
    $targetFolder = Get-FolderFromPath -FolderPath $targetFolderPath -MailboxName $TargetMailboxSmtp -Service $TargetService

    if ($DoNotCopyItems) {
        $result.Status = 'FolderOnly'
        return $result
    }

    if (-not $targetFolder) {
        $result.Status = 'TargetFolderNotFound'
        $result.Error  = ("Target folder '{0}' in mailbox '{1}' not found after creation." -f $targetFolderPath, $TargetMailboxSmtp)
        Write-Error $result.Error
        return $result
    }

    # Paging
    $iv = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    $iv.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    # Copy log
    $copyLog = ".\{0}.csv" -f ($TargetMailboxSmtp -replace '[^a-zA-Z0-9._-]','_')
    $already = @()
    if (Test-Path $copyLog) {
        $already = Import-Csv $copyLog -Encoding UTF8 | Select-Object -ExpandProperty UniqueId
        Write-Verbose "Previous copy results imported"
    } else {
        Write-Verbose "No file with previous copy results found"
    }

    $idx = 0; $itemTotal = 0
    do {
        # --- SOURCE read (PF service) ---
        try {
            $fi = $Service.FindItems($subFolderId,$iv)
        } catch {
            $m = $_.Exception.Message
            $hint = if ($m -match '(?i)401|Unauthorized|nicht autorisiert|AccessDenied|Zugriff verweigert') {
                " -> Check read permissions on public folder."
            } else { "" }
            $result.Status = 'SourceReadFailed'
            $result.Error  = $m + $hint
            Write-Error ("Items from source PF '{0}' could not be read: {1}{2}" -f $PublicFolderPath, $m, $hint)
            Write-Progress -Activity ("Copying items to MBX [{0}]" -f $TargetMailboxSmtp) -Completed
            return $result
        }

        $itemTotal = $fi.Items.Count
        $idx = 0

        foreach ($item in $fi.Items) {
            $subject = if ($item.Subject) { $item.Subject } else { '#no subject#' }
            $idx++
            $percent = if ($itemTotal -gt 0) { [int](($idx / $itemTotal) * 100) } else { 100 }
            Write-Progress -Activity ("{0}/{1} - Copying items to [{2}]" -f $idx,$itemTotal,$targetFolderPath) -Status $subject -PercentComplete $percent

            $uid = $item.Id.UniqueId
            if ($already -contains $uid) {
                $result.ItemsSkipped++
            } else {
                if ($PSCmdlet.ShouldProcess(("Item '{0}'" -f $subject), ("Copy to {0}" -f $targetFolder.DisplayName))) {
                    # Secure individual item: one error must not kill entire migration.
                    try {
                        # [void]: Item.Copy() returns copied item – otherwise it leaks
                        # into output stream and corrupts $migrationResults
                        [void]$item.Copy($targetFolder.Id)
                        [pscustomobject]@{ UniqueId = $uid } |
                            Export-Csv $copyLog -NoTypeInformation -Encoding UTF8 -Append
                        $result.ItemsCopied++
                    } catch {
                        $result.ItemsFailed++
                        Write-Warning ("Item '{0}' (UID {1}) could not be copied: {2}" -f $subject, $uid, $_.Exception.Message)
                    }
                }
            }
        }

        $iv.Offset += $fi.Items.Count
    } while ($fi.MoreAvailable -eq $true)

    Write-Progress -Activity ("Copying items to [{0}]" -f $targetFolderPath) -Completed

    if ($result.ItemsFailed -gt 0) {
        $result.Status = 'PartiallyFailed'
    }

    # Per-folder output is generated by the orchestrator (centralized, uniform format)
    return $result
}