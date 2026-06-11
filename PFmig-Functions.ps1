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

    # Wenn bereits eine EWS-Assembly im AppDomain ist, NICHT erneut laden
    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() |
              Where-Object { $_.GetName().Name -eq 'Microsoft.Exchange.WebServices' }
    if ($loaded) {
        Write-Verbose ("EWS bereits geladen: {0}" -f ($loaded | Select-Object -First 1 -Expand Location))
        return
    }

    # Sonst: höchste Version aus Registry laden (wie bisher)
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
    Write-Error "EWS Managed API (>=1.2) nicht gefunden. Bitte installieren."
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
        # 1) Autodiscover mit WIA versuchen
        try {
            $service.AutodiscoverUrl($MailboxName, { $true })
        } catch {
            # 2) Fallback: Skript laeuft auf dem Exchange-Server selbst.
            #    Autodiscover scheitert oft (kein SCP, Loopback, kein DNS-Record).
            #    Daher direkt den lokalen EWS-Endpunkt verwenden. localhost zuerst,
            #    weil das den Kerberos-Loopback-401 vermeidet.
            Write-Warning ("Autodiscover fehlgeschlagen ({0}). Versuche lokale EWS-Endpunkte." -f $_.Exception.Message)

            $fqdn = $null
            try { $fqdn = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName } catch {}

            $candidates = New-Object System.Collections.Generic.List[string]
            $candidates.Add("https://localhost/EWS/Exchange.asmx")
            if ($fqdn) { $candidates.Add(("https://{0}/EWS/Exchange.asmx" -f $fqdn)) }

            $connected = $false
            foreach ($candidate in $candidates) {
                try {
                    $service.Url = [System.Uri]$candidate
                    # Guenstige Validierung: Root-Ordner der Admin-Mailbox binden
                    $testId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
                        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName
                    )
                    [void][Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $testId)
                    Write-Host ("Mit EWS verbunden ueber Fallback-URL [{0}]" -f $candidate) -ForegroundColor DarkGray
                    $connected = $true
                    break
                } catch {
                    Write-Verbose ("Fallback-URL [{0}] fehlgeschlagen: {1}" -f $candidate, $_.Exception.Message)
                    $service.Url = $null
                }
            }
            if (-not $connected) {
                throw ("Verbindung zu EWS fehlgeschlagen. Bitte -Url explizit angeben (z. B. https://localhost/EWS/Exchange.asmx). Mailbox: '{0}'" -f $MailboxName)
            }
        }
    }

    if (-not $service.Url) { throw ("Error connecting to EWS (Mailbox '{0}')." -f $MailboxName) }
    return $service
}

function Get-TargetOU {
    <#
      Ermittelt die Ziel-OU (kompletter DN) fuer die Mailbox-Erstellung.
      Logik:
        - Existiert die Config-Datei mit einem gespeicherten DN -> fragen,
          ob dieser fuer den aktuellen Lauf verwendet werden soll (J/N).
        - Bei "Nein" oder fehlender Config -> Out-GridView-Auswahl.
        - Die getroffene OGV-Auswahl wird (ueber)in die Config-Datei geschrieben.
      Rueckgabe: DN-String oder $null bei Abbruch.
    #>
    [CmdletBinding()]
    param(
        [Parameter()][string]$ConfigPath,
        [Parameter()][string]$Title = "Mailbox creation: Select OU for disabled user account"
    )

    # 1) Gespeicherten DN lesen (falls Config vorhanden)
    $savedDn = $null
    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        try {
            $raw = (Get-Content -Path $ConfigPath -Raw -ErrorAction Stop).Trim()
            if (-not [string]::IsNullOrWhiteSpace($raw)) { $savedDn = $raw }
        } catch {
            Write-Warning ("Konnte OU-Config '{0}' nicht lesen: {1}" -f $ConfigPath, $_.Exception.Message)
        }
    }

    # 2) Gespeicherten DN anbieten
    if ($savedDn) {
        Write-Host ("Gespeicherte Ziel-OU gefunden:`n    {0}" -f $savedDn) -ForegroundColor Cyan
        $answer = Read-Host "Diese OU fuer den aktuellen Lauf verwenden? [J/N]"
        if ($answer -match '^(?i)\s*(j|ja|y|yes)\s*$') {
            return $savedDn
        }
        Write-Host "OK - neue OU-Auswahl per Out-GridView..." -ForegroundColor Yellow
    }

    # 3) OGV-Auswahl
    $ou = Get-ADOrganizationalUnit -Filter * | Select-Object DistinguishedName, Name |
          Out-GridView -PassThru -Title $Title
    if (-not $ou) { return $null }
    if ($ou -is [System.Array]) { $ou = $ou | Select-Object -First 1 }
    $ouDn = $ou.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($ouDn)) { return $null }

    # 4) Auswahl persistieren
    if ($ConfigPath) {
        try {
            Set-Content -Path $ConfigPath -Value $ouDn -Encoding UTF8 -ErrorAction Stop
            Write-Host ("Ziel-OU gespeichert in [{0}]" -f $ConfigPath) -ForegroundColor DarkGray
        } catch {
            Write-Warning ("Konnte OU nicht in Config '{0}' speichern: {1}" -f $ConfigPath, $_.Exception.Message)
        }
    }
    return $ouDn
}

function Test-EwsMailboxAccess {
    <#
      Prueft aktiv, ob ueber den uebergebenen EWS-Service auf die Mailbox
      zugegriffen werden kann (Bind auf MsgFolderRoot). Mit Retry, weil:
        - frisch erstellte Mailboxen im Store erst provisioniert werden muessen
        - die RBAC-/Impersonation-Konfiguration ein paar Minuten zum Greifen braucht
      Unterscheidet TRANSIENTE Fehler (warten + erneut versuchen) von
      PERMANENTEN Auth-/Impersonation-Fehlern (sofort sprechend abbrechen).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Parameter(Mandatory=$true)][string]$MailboxSmtp,
        [Parameter()][int]$MaxAttempts = 12,
        [Parameter()][int]$DelaySeconds = 10,
        [Parameter()][string]$Context = 'Ziel-Mailbox'
    )

    $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxSmtp
    )

    $lastMsg = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            [void][Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $rootId)
            Write-Verbose ("EWS-Zugriff auf {0} [{1}] bestaetigt (Versuch {2})." -f $Context, $MailboxSmtp, $attempt)
            return $true
        } catch {
            $msg = $_.Exception.Message
            $lastMsg = $msg
            # Permanente Auth-/Impersonation-Fehler erkennen -> nicht endlos retryen
            $isAuth = $msg -match '(?i)401|Unauthorized|nicht autorisiert|ImpersonateUserDenied|ImpersonationFailed|ImpersonateUser|AccessDenied|Zugriff verweigert'
            if ($isAuth) {
                throw ("EWS-Zugriff auf {0} [{1}] verweigert (Authentifizierung/Impersonation). " -f $Context, $MailboxSmtp) +
                      "Bitte pruefen: ApplicationImpersonation-Rolle fuer das ausfuehrende Konto zugewiesen? " +
                      "(New-ManagementRoleAssignment -Role ApplicationImpersonation -User <admin>) " +
                      "RBAC-Cache kann ~15 min brauchen. Originalfehler: $msg"
            }
            # Transient (Mailbox noch nicht provisioniert / Store kurz nicht erreichbar).
            # Eine dezente Wartemeldung beim ersten Versuch, Details nur mit -Verbose.
            if ($attempt -eq 1) {
                Write-Host ("  Warte auf Provisionierung der {0} [{1}] (bis zu {2}s)..." -f $Context, $MailboxSmtp, ($MaxAttempts * $DelaySeconds)) -ForegroundColor DarkGray
            }
            Write-Verbose ("EWS-Zugriff noch nicht moeglich (Versuch {0}/{1}): {2}" -f $attempt, $MaxAttempts, $msg)
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    $secs = $MaxAttempts * $DelaySeconds
    if ($lastMsg -match '(?i)temporarily unavailable|MailboxStoreUnavailable|failed to get the correct properties|voruebergehend nicht verfuegbar') {
        # Store/Datenbank meldet sich als nicht verfuegbar -> kein Impersonation-Problem.
        throw ("EWS-Zugriff auf {0} [{1}] nach {2}s nicht moeglich: der Mailbox-Store/die Datenbank ist 'temporarily unavailable'. " -f $Context, $MailboxSmtp, $secs) +
              "Das ist KEIN Berechtigungs-/Impersonationsproblem. Bitte pruefen: " +
              "Datenbank gemountet? (Get-MailboxDatabase -Status | ft Name,Mounted)  " +
              "Dienst 'Microsoft Exchange Information Store' laeuft? (Get-Service MSExchangeIS)  " +
              "Mailbox im Store sichtbar? (Get-MailboxStatistics -Identity '$MailboxSmtp')  " +
              "Die Mailbox existiert nun bereits - ein erneuter Start ueberspringt die Anlage und versucht den Zugriff direkt. " +
              "Hinweis: Laeuft Exchange auf einem Domain Controller, sind solche Store-Aussetzer haeufig (nicht unterstuetztes Setup). " +
              "Letzter Fehler: $lastMsg"
    }
    throw ("EWS-Zugriff auf {0} [{1}] nach {2}s nicht moeglich. " -f $Context, $MailboxSmtp, $secs) +
          "Mailbox evtl. noch nicht vollstaendig provisioniert, oder Berechtigung/Impersonation fehlt. Letzter Fehler: $lastMsg"
}

function Resolve-AutodiscoverUrl {
    <#
      Leitet die Autodiscover-.svc-URL ab. Reihenfolge:
        1) explizit uebergebener Wert
        2) aus dem (funktionierenden) EWS-Service-Host -> derselbe CAS bedient
           EWS und Autodiscover, daher zuverlaessig wenn EWS bereits verbunden ist
        3) bei localhost/127.0.0.1: stattdessen FQDN des Rechners (vermeidet
           Kerberos-Loopback-401 beim Autodiscover-SOAP-Call)
      Rueckgabe: URL-String oder $null.
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

    # Mehrere Strategien nacheinander versuchen (selbstheilend, falls eine scheitert):
    #   1) explizite/abgeleitete Autodiscover-URL (kein SCP)
    #   2) SCP-Lookup im AD (auf domaenengebundenem Exchange-Server sehr zuverlaessig)
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
            Write-Verbose ("Autodiscover-Versuch (Url={0}, Scp={1}) fehlgeschlagen: {2}" -f $s.Url, $s.Scp, $lastErr)
        }
    }

    if ($pfi) {
        if (-not $Service.HttpHeaders.$Header) { $Service.HttpHeaders.Add($Header, $pfi) }
    } else {
        throw ("Konnte PublicFolderInformation (X-AnchorMailbox) per Autodiscover nicht ermitteln. " +
               "Bitte -AutodiscoverUrl explizit angeben (z. B. https://ex.test.zarenko.net/autodiscover/autodiscover.svc). " +
               "Letzter Fehler: $lastErr")
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

        # URL bevorzugt aus EWS-Host ableiten (zuverlaessig, sobald EWS verbunden ist).
        $resolvedUrl = Resolve-AutodiscoverUrl -Service $Service -AutodiscoverUrl $AutodiscoverUrl
        if ($resolvedUrl) {
            $ads.Url = [System.Uri]$resolvedUrl
        } else {
            # Letzter Ausweg: Autodiscover selbst ermitteln lassen (SCP).
            $ads.EnableScpLookup = $true
            $null = $ads.GetUserSettings($MailboxName, [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::AutoDiscoverSMTPAddress)
        }

        if (-not $ads.Url) {
            throw "Autodiscover URL konnte nicht ermittelt werden. Übergib -AutodiscoverUrl (z. B. https://ex.test.zarenko.net/autodiscover/autodiscover.svc)."
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
    # Explizite [string[]]-Typisierung + Cast pro Element: verhindert, dass PS 5.1
    # bei einelementigem Where-Object-Output einen Skalar/Boolean leaked.
    [string[]]$segments = $FolderPath.Trim('\').Split('\') | Where-Object { $_ -ne '' } | ForEach-Object { [string]$_ }

    foreach ($seg in $segments) {
        $fv = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        if ($PropertySet) { $fv.PropertySet = $PropertySet }
        $sf = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $seg
        )
        # Retry-Schleife: EWS sieht neu erstellte Ordner nicht immer sofort
        $res = $null
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            $res = $Service.FindFolders($tf.Id, $sf, $fv)
            if ($res.TotalCount -gt 0) { break }
            Write-Verbose ("Get-FolderFromPath: Segment '{0}' noch nicht sichtbar, warte 2s (Versuch {1}/5)..." -f $seg, $attempt)
            Start-Sleep -Seconds 2
        }
        if ($res.TotalCount -gt 0) {
            foreach ($f in $res.Folders) { $tf = $f }
        } else {
            Write-Warning ("Get-FolderFromPath: Segment '{0}' in Pfad '{1}' nicht gefunden." -f $seg, $FolderPath)
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
        # 1) Sicherstellen, dass der X-AnchorMailbox-Header gesetzt ist
        if (-not $Service.HttpHeaders['X-AnchorMailbox']) {
            Get-PublicFolderRoutingHeader -Service $Service -MailboxName $SmtpAddress -Header 'X-AnchorMailbox' -AutodiscoverUrl $AutodiscoverUrl
        }

        # 2) PublicFoldersRoot binden + PropertySet vorbereiten
        $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot
        )

        $ps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
        )

        # PR_REPLICA_LIST optional hinzufügen (kann bei EWS-DLL-Mischbetrieb fehlschlagen)
        $PR_REPLICA_LIST = $null
        try {
            $PR_REPLICA_LIST = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
                0x6698, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary
            )
            # Expliziter Cast – bei homogener Assembly problemlos, bei Mismatch fangen wir ab
            [void]$ps.Add([Microsoft.Exchange.WebServices.Data.PropertyDefinitionBase]$PR_REPLICA_LIST)
        } catch {
            Write-Verbose "Konnte PR_REPLICA_LIST nicht zum PropertySet hinzufügen (möglicher EWS-Typkonflikt). Fahre ohne fort."
            $PR_REPLICA_LIST = $null
        }

        $tf = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $folderId, $ps)

        # 3) Root-Fall: "\" sofort zurückgeben
        if ([string]::IsNullOrEmpty($FolderPath) -or $FolderPath -eq '\') {
            # Optional: Routing-Header aus REPLICA_LIST ableiten (falls verfügbar)
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
                    Write-Verbose "PR_REPLICA_LIST am Root nicht verfügbar – verwende nur X-AnchorMailbox."
                }
            }
            return $tf.Id.UniqueId.ToString()
        }

        # 4) Teilbaum durchsuchen – Pfad sicher splitten (siehe Get-FolderFromPath)
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
                Write-Error "Error Folder Not Found (Segment: '$seg' im Pfad '$FolderPath')"
                return $null
            }
        }

        # 5) Nach erfolgreicher Auflösung: Content-Routing-Header optional ergänzen
        if ($PR_REPLICA_LIST) {
            $val = $null
            if ($tf.TryGetProperty($PR_REPLICA_LIST, [ref]$val)) {
                $guid = [System.Text.Encoding]::ASCII.GetString($val, 0, 36)
                $addr = New-Object System.Net.Mail.MailAddress($Service.HttpHeaders['X-AnchorMailbox'])
                $pfHeader = $guid + '@' + $addr.Host
                Write-Verbose ("Target Public Folder Routing Information Header : {0}" -f $pfHeader)

                # Setzt X-AnchorMailbox/X-PublicFolderMailbox für Content-Routing (Autodiscover-XML)
                Get-PublicFolderContentRoutingHeader -Service $Service -MailboxName $SmtpAddress -PfAddress $pfHeader -AutodiscoverUrl $AutodiscoverUrl
            } else {
                Write-Verbose "PR_REPLICA_LIST am Zielordner nicht verfügbar – verwende nur X-AnchorMailbox."
            }
        } else {
            Write-Verbose "PR_REPLICA_LIST nicht gesetzt – verwende nur X-AnchorMailbox."
        }

        # 6) UniqueId als String zurückgeben
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

    # Root direkt binden, wenn Parent leer oder "\" ist
    if ([string]::IsNullOrEmpty($ParentFolder) -or $ParentFolder -eq '\') {
        $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxName
        )
        $EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$rootId)
    } else {
        # Elternpfad sauber auflösen
        $EWSParentFolder = Get-FolderFromPath -MailboxName $MailboxName -Service $Service -FolderPath $ParentFolder
        if (-not $EWSParentFolder) { throw "Parent folder '$ParentFolder' not found in target mailbox '$MailboxName'." }
    }

    # Existenz prüfen
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

    # Status zurueckgeben, damit Aufrufer protokollieren koennen, was neu angelegt wurde.
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

      WICHTIG: $Service wird ausschliesslich fuer Public-Folder-EWS-Zugriffe verwendet
               (bekommt X-AnchorMailbox / X-PublicFolderMailbox Routing-Header).
               $TargetService ist eine separate, saubere EWS-Verbindung fuer den
               Zugriff auf die Ziel-Mailbox (ohne PF-Routing-Header).
               Beide Instanzen MUESSEN getrennt erstellt werden (Connect-Exchange).
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        # EWS-Service fuer Public Folder Zugriff (bekommt PF-Routing-Header)
        [Parameter(Mandatory=$true)][Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        # EWS-Service fuer Ziel-Mailbox Zugriff (OHNE PF-Routing-Header!)
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
    # Pfad in der Ziel-Mailbox vorab berechnen (fuer Reporting auch bei Fehlern)
    if ([string]::IsNullOrEmpty($ParentPath) -or $ParentPath -eq '\') {
        $targetFolderPath = '\' + $FolderName
    } else {
        $targetFolderPath = $ParentPath.TrimEnd('\') + '\' + $FolderName
    }

    # Ergebnis-Objekt, das am Ende (auch bei Fehlern) zurueckgegeben wird,
    # damit der Orchestrator eine sprechende Endabrechnung erstellen kann.
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

    # --- QUELLE: Public-Folder-Ordner-Id ermitteln (PF-Service) ---
    try {
        Get-PublicFolderRoutingHeader -Service $Service -MailboxName $AdminMailboxSmtp -Header "X-AnchorMailbox" -AutodiscoverUrl $AutodiscoverUrl
        $fldId = PublicFolderIdFromPath -Service $Service -FolderPath $PublicFolderPath -SmtpAddress $AdminMailboxSmtp -AutodiscoverUrl $AutodiscoverUrl
        if (-not $fldId) { throw "PublicFolderIdFromPath lieferte keine Id." }
        $subFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($fldId)
    } catch {
        $m = $_.Exception.Message
        $hint = if ($m -match '(?i)401|Unauthorized|nicht autorisiert|AccessDenied|Zugriff verweigert') {
            " -> QUELL-Zugriff verweigert: hat das ausfuehrende Konto Leserechte auf den Public Folder?"
        } else { "" }
        $result.Status = 'QuelleFehlgeschlagen'
        $result.Error  = $m + $hint
        Write-Error ("Quell-PF '{0}' nicht zugreifbar: {1}{2}" -f $PublicFolderPath, $m, $hint)
        return $result
    }

    # --- ZIEL: Ordner anlegen (TargetService, KEIN PF-Routing) ---
    try {
        $cf = Create-Folder -MailboxName $TargetMailboxSmtp -Service $TargetService -NewFolderName $FolderName -ParentFolder $ParentPath -FolderClass $FolderClass
        if ($cf) { $result.FolderCreated = [bool]$cf.Created }
    } catch {
        $m = $_.Exception.Message
        $hint = if ($m -match '(?i)401|Unauthorized|nicht autorisiert|Impersonat|AccessDenied|Zugriff verweigert') {
            " -> ZIEL-Zugriff verweigert: ApplicationImpersonation-Rolle gesetzt? Mailbox provisioniert?"
        } else { "" }
        $result.Status = 'ZielordnerFehlgeschlagen'
        $result.Error  = $m + $hint
        Write-Error ("Zielordner '{0}' konnte nicht erstellt werden: {1}{2}" -f $targetFolderPath, $m, $hint)
        return $result
    }

    # Zielordner aufloesen
    $targetFolder = Get-FolderFromPath -FolderPath $targetFolderPath -MailboxName $TargetMailboxSmtp -Service $TargetService

    if ($DoNotCopyItems) {
        $result.Status = 'NurOrdner'
        return $result
    }

    if (-not $targetFolder) {
        $result.Status = 'ZielordnerNichtGefunden'
        $result.Error  = ("Zielordner '{0}' in Mailbox '{1}' nach Anlegen nicht auffindbar." -f $targetFolderPath, $TargetMailboxSmtp)
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
        # --- QUELLE lesen (PF-Service) ---
        try {
            $fi = $Service.FindItems($subFolderId,$iv)
        } catch {
            $m = $_.Exception.Message
            $hint = if ($m -match '(?i)401|Unauthorized|nicht autorisiert|AccessDenied|Zugriff verweigert') {
                " -> QUELL-Leserechte auf den Public Folder pruefen."
            } else { "" }
            $result.Status = 'QuelleLesenFehlgeschlagen'
            $result.Error  = $m + $hint
            Write-Error ("Items aus Quell-PF '{0}' konnten nicht gelesen werden: {1}{2}" -f $PublicFolderPath, $m, $hint)
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
                    # Einzelnes Item absichern: ein Fehler darf nicht die ganze Migration killen.
                    try {
                        # [void]: Item.Copy() gibt das kopierte Item zurueck - sonst leakt
                        # es in den Output-Stream der Funktion und verfaelscht $migrationResults.
                        [void]$item.Copy($targetFolder.Id)
                        [pscustomobject]@{ UniqueId = $uid } |
                            Export-Csv $copyLog -NoTypeInformation -Encoding UTF8 -Append
                        $result.ItemsCopied++
                    } catch {
                        $result.ItemsFailed++
                        Write-Warning ("Item '{0}' (UID {1}) konnte nicht kopiert werden: {2}" -f $subject, $uid, $_.Exception.Message)
                    }
                }
            }
        }

        $iv.Offset += $fi.Items.Count
    } while ($fi.MoreAvailable -eq $true)

    Write-Progress -Activity ("Copying items to [{0}]" -f $targetFolderPath) -Completed

    if ($result.ItemsFailed -gt 0) {
        $result.Status = 'TeilweiseFehlgeschlagen'
    }

    # Per-Ordner-Ausgabe macht der Orchestrator (zentral, einheitliches Format).
    return $result
}