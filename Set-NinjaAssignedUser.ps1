<#
.SYNOPSIS
    Bulk-assigns the NinjaOne "Assigned User" (device owner) from a CSV.

.DESCRIPTION
    Reads a device/user CSV, resolves each row to exactly one device and one owner,
    validates the organization rule, then sets the Assigned User via
    POST /v2/device/{id}/owner/{ownerUid}.

    Runs in DRY RUN mode by default. Nothing is written until you pass -Execute.

    -- RUN AS A NINJAONE AUTOMATION ---------------------------------------------
    1. API app: Administration > Apps > API > Client app IDs > Add.
       Platform API Services, grant type Client credentials, scopes Monitoring
       and Management.
    2. Device custom fields, type Secure, Automations set to Read Only:
           ninjaoneClientId       the Client ID
           ninjaoneClientSecret   the Client Secret
       Fill both in on the device that will run this script.
    3. Save the CSV on that device as C:\Temp\Assigned-Users.csv
    4. Add this script to the Automation library: PowerShell, Windows, Run As System.
    5. Run it on that device:
           Dry run : no parameters
           Apply   : -Execute
       The report is written next to the CSV. Read it before applying.

    -- CSV COLUMNS --------------------------------------------------------------
    Headers are matched case-insensitively and ignore spaces/underscores. Recognised:

        Device      : Agent Name, Device, DeviceName, Hostname, SystemName
        Organization: Client Name, Organization, Org, Client   (optional)
        Username    : Agent User, Username, SamAccountName     (optional)
        Last login  : LastLogin, LastLoggedOnUser, LoggedOnUser (optional)
        OS          : Agent Operating System, OperatingSystem, OS (optional)
        Email       : Email, EmailAddress, UserEmail, Mail     (optional)
        First name  : FirstName, GivenName                     (optional)
        Last name   : LastName, Surname, FamilyName            (optional)
        Full name   : Name, FullName, DisplayName              (optional)

    Only the device column is strictly required. Rows are handled like this:

        DOMAIN\First.Last     Looked up as the person First Last.
        Not Logged In         Reported as NOUSER. The device is left alone.
        admin, patchmgr, ...  Reported as FILTERED. See -ExcludeUserPattern.
        Windows Server        Reported as FILTERED unless -IncludeServers is set.
        #N/A and other Excel errors count as blank and are reported, never guessed at.

    -- OWNER RESOLUTION ---------------------------------------------------------
    Owners come from Contacts, End users, and Technicians (with -IncludeTechnicians).

    1. A row with an email matches on the full address.
    2. A row with no email column matches on the person's name, or on the part of
       the owner's email before the @. Case, dots and accents are ignored.

    When a name exists more than once, the record in the device's organization wins,
    then a record with no organization. Anything still tied is an error, not a guess.

    -- ORGANIZATION RULE (verified against the live API) ------------------------
    The owner's organizationId must be absent (global record) or exactly equal to
    the device's organizationId. Anything else returns
    HTTP 403 owner_has_not_same_node_organization.

.PARAMETER CsvPath
    Path to the input CSV. Defaults to C:\Temp\Assigned-Users.csv

.PARAMETER Instance
    NinjaOne instance hostname. Defaults to app.ninjarmm.com (US). Others include
    app.ninjaone.com, eu.ninjaone.com, ca.ninjarmm.com. A full base URL also works.

.PARAMETER Execute
    Perform writes. Without this switch the script only produces a dry-run report.

.PARAMETER ClientIdField
    Custom field holding the API Client ID. Defaults to ninjaoneClientId.

.PARAMETER ClientSecretField
    Custom field holding the API Client Secret. Defaults to ninjaoneClientSecret.

.PARAMETER IncludeServers
    Also process rows whose OS column says Server. They are skipped by default.

.PARAMETER EmailDomain
    Build a missing email from the person's name, e.g. -EmailDomain example.com
    produces Benjamin.Ong@example.com. Only used when the Email column is blank.

.PARAMETER EmailPattern
    How to build a derived email. first.last (default), firstlast, flast,
    first_last, or username (uses the sAMAccountName).

.PARAMETER LowercaseEmail
    Force derived emails to lowercase.

.PARAMETER MatchByName
    When a sheet does have an email column, fall back to matching on name for rows
    whose email is blank or unknown. Always on for a sheet with no email source.

.PARAMETER CreateMissingContacts
    Create a contact when no existing contact/end user matches. The contact is
    created in the device's own organization. Requires a resolvable email.

.PARAMETER IncludeTechnicians
    Also consider technician records when resolving an owner.

.PARAMETER ExcludeUserPattern
    Extra regex of usernames/emails to skip, added to the built-in list of local
    accounts (admin, administrator, patchmgr, localadmin, guest, test, user).
    Example: -ExcludeUserPattern 'adm$|Act\d+$|^svc'

.PARAMETER OnlyOrganization
    Process only rows whose Organization column matches this name. Useful for
    piloting one org out of a large export.

.PARAMETER LogPath
    Path for the CSV result log. Defaults to ninja-assigned-user-<timestamp>.csv
    next to the input CSV.

.PARAMETER Force
    Skip the confirmation prompt in an interactive session. An automation run has
    no prompt: there, -Execute is the confirmation.

.EXAMPLE
    # As an automation, in the Parameters box of the Run dialog: apply the changes
    -Execute

.EXAMPLE
    # Pilot one organization
    -OnlyOrganization "North Region" -Execute

.EXAMPLE
    # From a PowerShell window, with the credentials in environment variables
    .\Set-NinjaAssignedUser.ps1 -CsvPath .\Assigned-Users.csv

.NOTES
    Version 2.2

    Credentials are read from the two custom fields above when the script runs as
    an automation, otherwise from $env:NINJA_CLIENT_ID and $env:NINJA_CLIENT_SECRET.
    Nothing is hardcoded and no token or secret is written to the output or the log.

    Required OAuth scopes    : "monitoring management"
    Required API app settings: Client Credentials grant, platform = API Services.

    Rollback: the log's PreviousOwnerUid column holds the prior value. Clear an
    assignment with DELETE /v2/device/{deviceId}/owner
#>

[CmdletBinding()]
param(
    [string] $CsvPath = 'C:\Temp\Assigned-Users.csv',
    [string] $Instance = 'app.ninjarmm.com',
    [switch] $Execute,
    [string] $ClientIdField = 'ninjaoneClientId',
    [string] $ClientSecretField = 'ninjaoneClientSecret',
    [switch] $IncludeServers,
    [string] $EmailDomain,
    [ValidateSet('first.last','firstlast','flast','first_last','username')]
    [string] $EmailPattern = 'first.last',
    [switch] $LowercaseEmail,
    [switch] $MatchByName,
    [switch] $CreateMissingContacts,
    [switch] $IncludeTechnicians,
    [string] $ExcludeUserPattern,
    [string] $OnlyOrganization,
    [string] $LogPath,
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# PowerShell 5.1 defaults to TLS 1.0, which NinjaOne rejects.
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

# -----------------------------------------------------------------------------
# Script state
# -----------------------------------------------------------------------------
$script:Token    = $null
$script:TokenExp = [datetime]::MinValue
$script:Results  = [System.Collections.Generic.List[object]]::new()

if ($Instance -match '^https?://') {
    $script:BaseUrl = $Instance.TrimEnd('/')
} else {
    $script:BaseUrl = "https://$($Instance.TrimEnd('/'))"
}

# Row-by-row console output is only useful on a short sheet. A long one would bury
# the summary in the automation's activity output; the report file has every row.
$script:MaxRowsToPrint = 50

function Write-Step { param([string]$Message) Write-Host "`n=== $Message ===" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Message) Write-Host "  [ OK ] $Message" -ForegroundColor Green }
function Write-Warn { param([string]$Message) Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
function Write-Err  { param([string]$Message) Write-Host "  [FAIL] $Message" -ForegroundColor Red }
function Write-Skip { param([string]$Message) Write-Host "  [SKIP] $Message" -ForegroundColor DarkGray }

# -----------------------------------------------------------------------------
# CSV hygiene
# -----------------------------------------------------------------------------

# Excel formula errors arrive as literal text. They are absence of data, not data.
$script:ExcelErrors = @('#VALUE!','#N/A','#NA','#REF!','#NAME?','#DIV/0!','#NULL!','#NUM!','#SPILL!','#CALC!','#GETTING_DATA')

# What an RMM export writes when nobody was signed in. A real answer, not a data gap.
$script:NoUserMarkers = @('not logged in','not logged on','no user','no user logged in','nobody')

# Local accounts that are never a person, compared against the bare account name.
$script:LocalAccountPattern = '^(admin|administrator|patchmgr|localadmin|guest|test|user)$'

function Clean-Value {
    <# Trims, and collapses Excel errors / placeholders to an empty string. #>
    param([object] $Value)

    if ($null -eq $Value) { return '' }
    $s = [string]$Value
    $s = $s.Trim().Trim([char]0x00A0)      # non-breaking space survives copy/paste
    if ($s -eq '') { return '' }
    if ($script:ExcelErrors -contains $s.ToUpperInvariant()) { return '' }
    if ($s -in @('-','--','N/A','NA','n/a','null','NULL','none','None')) { return '' }
    return $s
}

function Test-NoUserMarker {
    param([string] $Value)
    return ($script:NoUserMarkers -contains $Value.Trim().ToLowerInvariant())
}

function Get-CsvDelimiter {
    <# Cells pasted from Excel into Notepad arrive tab-separated, not comma-separated. #>
    param([Parameter(Mandatory)][string] $Path)

    $line   = [string](Get-Content -LiteralPath $Path -TotalCount 1)
    $commas = ($line -split ',').Count
    $tabs   = ($line -split "`t").Count
    $semis  = ($line -split ';').Count

    if ($tabs -gt $commas -and $tabs -ge $semis) { return "`t" }
    if ($semis -gt $commas) { return ';' }
    return ','
}

function Get-ColumnMap {
    <#
        Maps our logical field names onto whatever the export actually calls them.
        Comparison ignores case, spaces, underscores and hyphens.
    #>
    param([Parameter(Mandatory)][string[]] $Headers)

    $aliases = [ordered]@{
        Device       = @('agentname','device','devicename','hostname','host','systemname','computer','computername','machine','machinename','nameofdevice')
        Organization = @('clientname','organization','organisation','org','organizationname','orgname','client','customer')
        Location     = @('location','locationname','site')
        Email        = @('email','emailaddress','mail','useremail','primaryemail','upn','userprincipalname')
        FirstName    = @('firstname','first','givenname','forename')
        LastName     = @('lastname','last','surname','familyname')
        FullName     = @('name','fullname','displayname','username_display','user','person','assigneduser','owner')
        Username     = @('agentuser','username','samaccountname','samaccount','logon','login','account','accountname','userid')
        LastLogin    = @('lastlogin','lastloggedonuser','loggedonuser','lastloggedinuser','lastuser','lastlogon')
        OS           = @('agentoperatingsystem','operatingsystem','os','osname')
    }

    $norm = @{}
    foreach ($h in $Headers) {
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $key = ($h -replace '[\s_\-\.]', '').ToLowerInvariant()
        if (-not $norm.ContainsKey($key)) { $norm[$key] = $h }
    }

    $map = @{}
    foreach ($field in $aliases.Keys) {
        foreach ($alias in $aliases[$field]) {
            if ($norm.ContainsKey($alias)) { $map[$field] = $norm[$alias]; break }
        }
    }

    # 'Name'/'User' are ambiguous: they can land on both FullName and Username.
    # If both resolved to the same header, keep it as FullName only.
    if ($map.ContainsKey('FullName') -and $map.ContainsKey('Username') -and
        $map['FullName'] -eq $map['Username']) {
        $map.Remove('Username')
    }
    return $map
}

function Get-Field {
    <# Reads a logical field from a CSV row via the column map, cleaned. #>
    param(
        [Parameter(Mandatory)] $Row,
        [Parameter(Mandatory)][hashtable] $Map,
        [Parameter(Mandatory)][string] $Field
    )
    if (-not $Map.ContainsKey($Field)) { return '' }
    $header = $Map[$Field]
    if ($Row.PSObject.Properties.Name -notcontains $header) { return '' }
    return (Clean-Value $Row.$header)
}

function Split-PersonName {
    <#
        "Benjamin Ong"              -> Benjamin / Ong
        "Srinivasa Kalagana ADM"    -> Srinivasa / Kalagana ADM
        "Ong, Benjamin"             -> Benjamin / Ong
        "Cher"                      -> Cher / Cher   (API needs both fields)
    #>
    param([string] $FullName)

    $n = (Clean-Value $FullName)
    if ($n -eq '') { return $null }

    if ($n -match '^\s*([^,]+)\s*,\s*(.+)$') {
        return [pscustomobject]@{ First = $Matches[2].Trim(); Last = $Matches[1].Trim() }
    }

    $parts = @($n -split '\s+' | Where-Object { $_ -ne '' })
    if ($parts.Count -eq 0) { return $null }
    if ($parts.Count -eq 1) {
        return [pscustomobject]@{ First = $parts[0]; Last = $parts[0] }
    }
    return [pscustomobject]@{ First = $parts[0]; Last = ($parts[1..($parts.Count - 1)] -join ' ') }
}

function Get-BareUsername {
    <# "AP\bw6235" / "NA\GBDSA.NOCAct2" / "user@corp.com" -> the account name. #>
    param([string] $Value)

    $v = (Clean-Value $Value)
    if ($v -eq '') { return '' }
    if ($v -match '\\') { $v = ($v -split '\\')[-1] }
    if ($v -match '/')  { $v = ($v -split '/')[-1]  }
    if ($v -match '^([^@]+)@') { $v = $Matches[1] }
    return $v.Trim()
}

function Get-IdentityKey {
    <# "Dianet.Aleman", "dianet aleman" and "DIANET-ALEMAN" all reduce to one key. #>
    param([string] $Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $plain = $Value.Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
    return ($plain -replace '[^\p{L}\p{N}]', '').ToLowerInvariant()
}

function New-DerivedEmail {
    param([string] $First, [string] $Last, [string] $Username, [string] $Domain, [string] $Pattern)

    if ([string]::IsNullOrWhiteSpace($Domain)) { return '' }

    # Strip characters that are not legal in the local part of an address.
    $f = ($First -replace '[^A-Za-z0-9''\-]', '')
    $l = ($Last  -replace '[^A-Za-z0-9''\-]', '')
    $u = ($Username -replace '[^A-Za-z0-9''\-\.]', '')

    $local = ''
    switch ($Pattern) {
        'username'   { $local = $u }
        'firstlast'  { if ($f -and $l) { $local = "$f$l" } }
        'flast'      { if ($f -and $l) { $local = "$($f.Substring(0,1))$l" } }
        'first_last' { if ($f -and $l) { $local = "${f}_${l}" } }
        default      { if ($f -and $l) { $local = "$f.$l" } }   # first.last
    }

    if ([string]::IsNullOrWhiteSpace($local)) { return '' }
    $addr = "$local@$($Domain.TrimStart('@'))"
    if ($LowercaseEmail) { $addr = $addr.ToLowerInvariant() }
    return $addr
}

# -----------------------------------------------------------------------------
# Auth - client credentials. Token is cached and refreshed 60s before expiry.
# -----------------------------------------------------------------------------
function Get-FieldCredential {
    <#
        Reads a credential from a device custom field. Returns '' when it cannot.
        The value is never printed: a failed read returns text, not an exception, so
        anything with whitespace in it is an error message and not a credential.
    #>
    param([Parameter(Mandatory)][string] $FieldName)

    if (-not (Get-Command Ninja-Property-Get -ErrorAction SilentlyContinue)) { return '' }

    try {
        $output = @(Ninja-Property-Get $FieldName 2>$null | ForEach-Object { [string]$_ })
    } catch {
        return ''
    }
    $value = ($output -join '').Trim()
    if ($value -match '\s') { return '' }
    # Copying from a web page or chat can bring quotes and invisible characters along.
    return ($value -replace '[^\x21-\x7E]', '').Trim('"', "'")
}

function Get-ApiCredentials {
    $id         = Get-FieldCredential -FieldName $ClientIdField
    $secret     = Get-FieldCredential -FieldName $ClientSecretField
    $source     = "custom fields $ClientIdField / $ClientSecretField"
    $fromFields = $true

    if (-not $id -or -not $secret) {
        $id         = $env:NINJA_CLIENT_ID
        $secret     = $env:NINJA_CLIENT_SECRET
        $source     = 'environment variables'
        $fromFields = $false
    }

    if ([string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($secret)) {
        throw "No API credentials found. Fill in the custom fields $ClientIdField and $ClientSecretField " +
              "on this device, set their Automations permission to Read Only, and run this as an automation " +
              "with Run As System. Outside an automation, set NINJA_CLIENT_ID and NINJA_CLIENT_SECRET instead."
    }
    return [pscustomobject]@{ Id = $id; Secret = $secret; Source = $source; FromFields = $fromFields }
}

$script:Credentials = Get-ApiCredentials

function Get-NinjaToken {
    if ($script:Token -and (Get-Date) -lt $script:TokenExp) { return $script:Token }

    # Encoded by hand so PowerShell 5.1 and 7 send the same bytes, whatever is in the secret.
    $body = 'grant_type=client_credentials' +
            '&client_id='     + [uri]::EscapeDataString($script:Credentials.Id) +
            '&client_secret=' + [uri]::EscapeDataString($script:Credentials.Secret) +
            '&scope='         + [uri]::EscapeDataString('monitoring management')

    try {
        $resp = Invoke-RestMethod -Method Post -Uri "$script:BaseUrl/ws/oauth/token" `
                                  -Body $body -ContentType 'application/x-www-form-urlencoded'
    }
    catch {
        # Deliberately does NOT echo the body - it would contain the client secret.
        $status = 'no response'
        if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
            try { $status = "HTTP $([int]$_.Exception.Response.StatusCode)" } catch { }
        }
        $said = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $said = ($_.ErrorDetails.Message -replace '\s+', ' ').Trim() }

        $meaning = 'Check the instance hostname and that this device can reach it.'
        if ($status -eq 'HTTP 404')          { $meaning = 'NinjaOne does not recognise this Client ID on this instance.' }
        elseif ($said -match 'invalid_client') { $meaning = 'NinjaOne knows this Client ID but rejected the Client Secret. Copy the secret again, or generate a new one.' }
        elseif ($said -match 'invalid_scope')  { $meaning = 'The API app does not allow the scopes this script asks for. Tick Monitoring and Management on the app.' }
        elseif ($said -match 'grant')          { $meaning = 'The API app does not allow this login type. Tick Client credentials on the app.' }

        $shownId = $script:Credentials.Id
        if ($shownId.Length -gt 10) { $shownId = "$($shownId.Substring(0,6))...$($shownId.Substring($shownId.Length - 4))" }

        Write-Host ""
        Write-Err "Login to $script:BaseUrl failed ($status)"
        if ($said) { Write-Host "  NinjaOne said : $said" }
        Write-Host "  What it means : $meaning"
        Write-Host "  Client ID     : $shownId ($($script:Credentials.Id.Length) characters)"
        Write-Host "  Client Secret : $($script:Credentials.Secret.Length) characters"
        Write-Host "  Read from     : $($script:Credentials.Source)"
        Write-Host ""
        throw "Login failed ($status). $meaning"
    }

    $script:Token    = $resp.access_token
    $expiresIn       = if ($resp.PSObject.Properties.Name -contains 'expires_in') { [int]$resp.expires_in } else { 3600 }
    $script:TokenExp = (Get-Date).AddSeconds($expiresIn - 60)
    return $script:Token
}

# -----------------------------------------------------------------------------
# HTTP wrapper with retry/backoff. Returns a result object rather than throwing,
# so a single bad row can never abort the whole run.
# -----------------------------------------------------------------------------
function Invoke-NinjaApi {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PATCH','PUT','DELETE')][string] $Method,
        [Parameter(Mandatory)][string] $Path,
        [object] $Body,
        [int]    $MaxAttempts = 4
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        # Outside the try: bad credentials should stop the run with their own message.
        $token = Get-NinjaToken

        try {
            $headers = @{ Authorization = "Bearer $token" }
            $params  = @{
                Method      = $Method
                Uri         = "$script:BaseUrl/v2$Path"
                Headers     = $headers
                ContentType = 'application/json'
            }
            if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress) }

            $data = Invoke-RestMethod @params
            return [pscustomobject]@{ Success = $true; Data = $data; Status = 200; Error = $null }
        }
        catch {
            $status = 0
            if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
            }

            # Read the response body for a usable error code
            # (e.g. owner_has_not_same_node_organization).
            # ErrorDetails is null when there was no response at all.
            $detail = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = $_.ErrorDetails.Message }

            # 401 -> token may have been revoked mid-run. Drop it and retry once.
            if ($status -eq 401 -and $attempt -lt $MaxAttempts) {
                $script:Token = $null; $script:TokenExp = [datetime]::MinValue
                Start-Sleep -Milliseconds 500
                continue
            }

            # Transient: rate limit or server-side. Exponential backoff.
            if (($status -eq 429 -or $status -ge 500) -and $attempt -lt $MaxAttempts) {
                $delay = [math]::Pow(2, $attempt)
                Write-Warn "HTTP $status on $Method $Path - retrying in ${delay}s (attempt $attempt/$MaxAttempts)"
                Start-Sleep -Seconds $delay
                continue
            }

            # Terminal for this call.
            return [pscustomobject]@{ Success = $false; Data = $null; Status = $status; Error = $detail }
        }
    }

    return [pscustomobject]@{ Success = $false; Data = $null; Status = 0; Error = 'Retry attempts exhausted' }
}

# -----------------------------------------------------------------------------
# Bulk reads
# -----------------------------------------------------------------------------
function Get-AllDevices {
    $all   = [System.Collections.Generic.List[object]]::new()
    $after = 0
    $page  = 1000

    while ($true) {
        $res = Invoke-NinjaApi -Method GET -Path "/devices?pageSize=$page&after=$after"
        if (-not $res.Success) { throw "Failed to list devices (HTTP $($res.Status)): $($res.Error)" }

        $batch = @($res.Data)
        if ($batch.Count -eq 0) { break }

        $all.AddRange($batch)
        if ($batch.Count -lt $page) { break }
        $after = $batch[-1].id
    }
    return $all
}

function Get-Organizations {
    <#
        ByName : lowercase org name -> id. Used to disambiguate duplicate hostnames.
        ById   : id -> org name. Used to say which org a person actually belongs to.
    #>
    $byName = @{}
    $byId   = @{}
    $res = Invoke-NinjaApi -Method GET -Path '/organizations?pageSize=1000'
    if (-not $res.Success) {
        Write-Warn "Could not read /organizations (HTTP $($res.Status)) - Organization column will be ignored."
    } else {
        foreach ($o in @($res.Data)) {
            if (-not $o.name) { continue }
            $byName[$o.name.Trim().ToLowerInvariant()] = $o.id
            $byId[[string]$o.id] = $o.name
        }
    }
    return [pscustomobject]@{ ByName = $byName; ById = $byId }
}

function Get-OwnerDirectory {
    <#
        Returns two indexes over the same owner records:
          ByEmail    : lowercase email -> owner
          ByIdentity : identity key    -> every owner with that name or email prefix
    #>
    $byEmail    = @{}
    $byIdentity = @{}

    $sources = @(
        @{ Path = '/contacts';       Type = 'Contact' },
        @{ Path = '/user/end-users'; Type = 'EndUser' }
    )
    if ($IncludeTechnicians) { $sources += @{ Path = '/user/technicians'; Type = 'Technician' } }

    foreach ($src in $sources) {
        $res = Invoke-NinjaApi -Method GET -Path $src.Path
        if (-not $res.Success) {
            Write-Warn "Could not read $($src.Path) (HTTP $($res.Status)) - skipping this owner source."
            continue
        }

        foreach ($rec in @($res.Data)) {
            $orgId = $null
            if ($rec.PSObject.Properties.Name -contains 'organizationId') { $orgId = $rec.organizationId }

            $first = ''
            if ($rec.PSObject.Properties.Name -contains 'firstName' -and $rec.firstName) { $first = [string]$rec.firstName }
            $last = ''
            if ($rec.PSObject.Properties.Name -contains 'lastName'  -and $rec.lastName)  { $last  = [string]$rec.lastName }

            $email = ''
            if ($rec.PSObject.Properties.Name -contains 'email' -and $rec.email) { $email = [string]$rec.email }

            $owner = [pscustomobject]@{
                Uid            = $rec.uid
                Id             = $rec.id
                FirstName      = $first
                LastName       = $last
                Email          = $email
                OrganizationId = $orgId
                Type           = $src.Type
            }

            if ($email) {
                $ekey = $email.Trim().ToLowerInvariant()
                # First source wins: Contact > EndUser > Technician.
                if (-not $byEmail.ContainsKey($ekey)) { $byEmail[$ekey] = $owner }
            }

            $identityKeys = @(
                (Get-IdentityKey "$first $last"),
                (Get-IdentityKey ($email -replace '@.*$', ''))
            ) | Where-Object { $_ } | Select-Object -Unique

            foreach ($ikey in $identityKeys) {
                if (-not $byIdentity.ContainsKey($ikey)) {
                    $byIdentity[$ikey] = [System.Collections.Generic.List[object]]::new()
                }
                $byIdentity[$ikey].Add($owner)
            }
        }
    }

    return [pscustomobject]@{ ByEmail = $byEmail; ByIdentity = $byIdentity }
}

function Select-OwnerForDevice {
    <#
        Picks one owner out of the records sharing a name, using the org rule the API
        enforces. Returns the owner, or the reason there is not exactly one.
    #>
    param(
        [Parameter(Mandatory)][object[]] $Candidates,
        [Parameter(Mandatory)][int] $DeviceOrgId
    )

    # A person can be both a contact and an end user. Same email = same person.
    $unique = [System.Collections.Generic.List[object]]::new()
    $seen   = @{}
    foreach ($c in $Candidates) {
        $id = if ($c.Email) { $c.Email.ToLowerInvariant() } else { [string]$c.Uid }
        if (-not $seen.ContainsKey($id)) { $seen[$id] = $true; $unique.Add($c) }
    }

    $sameOrg = @($unique | Where-Object { $_.OrganizationId -eq $DeviceOrgId })
    $global  = @($unique | Where-Object { $null -eq $_.OrganizationId })

    foreach ($tier in @(,$sameOrg) + @(,$global)) {
        if ($tier.Count -eq 1) { return [pscustomobject]@{ Owner = $tier[0]; Problem = ''; Others = $null } }
        if ($tier.Count -gt 1) { return [pscustomobject]@{ Owner = $null; Problem = 'AMBIGUOUS'; Others = $tier } }
    }
    return [pscustomobject]@{ Owner = $null; Problem = 'OTHER_ORG'; Others = @($unique) }
}

function New-NinjaContact {
    param(
        [Parameter(Mandatory)][string] $FirstName,
        [Parameter(Mandatory)][string] $LastName,
        [Parameter(Mandatory)][string] $Email,      # required by the API - non-nullable
        [Parameter(Mandatory)][int]    $OrganizationId
    )

    $body = @{
        firstName      = $FirstName
        lastName       = $LastName
        email          = $Email
        organizationId = $OrganizationId
    }

    $res = Invoke-NinjaApi -Method POST -Path '/contacts' -Body $body

    # 409 = the email already exists somewhere in the tenant. Treat as recoverable.
    if (-not $res.Success -and $res.Status -eq 409) {
        return [pscustomobject]@{ Success = $false; AlreadyExists = $true; Data = $null; Error = 'EMAIL_ALREADY_IN_USE' }
    }
    return [pscustomobject]@{
        Success       = $res.Success
        AlreadyExists = $false
        Data          = $res.Data
        Error         = $res.Error
    }
}

# -----------------------------------------------------------------------------
# Load and map the CSV
# -----------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV not found: $CsvPath. Save the sheet there, or pass its location with -CsvPath."
}
$CsvPath = (Resolve-Path -LiteralPath $CsvPath).Path

# Excel (or OneDrive sync) can hold the file open. Open it with shared read/write so a
# lock doesn't stop the run, and work from a temp copy of what was read.
$CsvOriginal = $CsvPath
$CsvCopy     = Join-Path ([IO.Path]::GetTempPath()) ("ninja-assigned-" + [guid]::NewGuid().ToString('N') + '.csv')
try {
    $share = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
    $src   = New-Object IO.FileStream($CsvOriginal, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
    try {
        $dst = [IO.File]::Create($CsvCopy)
        try { $src.CopyTo($dst) } finally { $dst.Dispose() }
    } finally { $src.Dispose() }
}
catch {
    throw "Could not read '$CsvOriginal': $($_.Exception.Message). Close the file in Excel and run again."
}
$CsvPath = $CsvCopy

# An Excel workbook renamed to .csv is really a zip file, which starts with "PK".
$stream = [IO.File]::OpenRead($CsvPath)
try { $firstByte = $stream.ReadByte(); $secondByte = $stream.ReadByte() } finally { $stream.Dispose() }
if ($firstByte -eq 0x50 -and $secondByte -eq 0x4B) {
    Remove-Item -LiteralPath $CsvCopy -Force -ErrorAction SilentlyContinue
    throw "'$CsvOriginal' is an Excel workbook, not a CSV. Open it in Excel and use Save As with the type " +
          "CSV UTF-8 (Comma delimited). Renaming the file to .csv does not convert it."
}

if (-not $LogPath) {
    $stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
    $LogPath = Join-Path (Split-Path -Parent $CsvOriginal) "ninja-assigned-user-$stamp.csv"
}

$rows = @(Import-Csv -LiteralPath $CsvPath -Delimiter (Get-CsvDelimiter -Path $CsvPath))
if ($rows.Count -eq 0) { throw "CSV '$CsvPath' contains no data rows." }

$headers = @($rows[0].PSObject.Properties.Name)
$colMap  = Get-ColumnMap -Headers $headers

if (-not $colMap.ContainsKey('Device')) {
    throw "Could not find a device column in the CSV. Looked for: Agent Name, Device, DeviceName, Hostname, SystemName. " +
          "Found headers: $($headers -join ', ')"
}

$hasIdentitySource = $colMap.ContainsKey('Email') -or $colMap.ContainsKey('FullName') -or
                     $colMap.ContainsKey('FirstName') -or $colMap.ContainsKey('Username') -or
                     $colMap.ContainsKey('LastLogin')
if (-not $hasIdentitySource) {
    throw "The CSV has a device column but nothing to identify a person. Add at least one of: " +
          "Agent User, Email, Name, First Name/Last Name, Username, Last Login."
}

# With no way to get an address, the name is the only thing left to match on.
$useNameMatch = $MatchByName -or (-not $colMap.ContainsKey('Email') -and -not $EmailDomain)

Write-Host ""
Write-Host "NinjaOne - Bulk Assigned User (v2.2)" -ForegroundColor White
Write-Host "Instance : $script:BaseUrl"
Write-Host "CSV      : $CsvOriginal ($($rows.Count) rows)"
Remove-Item -LiteralPath $CsvCopy -Force -ErrorAction SilentlyContinue
Write-Host "Auth     : $($script:Credentials.Source)"
Write-Host "Mode     : $(if ($Execute) { 'EXECUTE (writes enabled)' } else { 'DRY RUN (no changes)' })" `
           -ForegroundColor $(if ($Execute) { 'Yellow' } else { 'Green' })

Write-Host "Columns  : " -NoNewline
Write-Host (($colMap.Keys | Sort-Object | ForEach-Object { "$_ <- '$($colMap[$_])'" }) -join ', ') -ForegroundColor DarkGray
if ($EmailDomain)  { Write-Host "Email    : derive missing as $EmailPattern@$($EmailDomain.TrimStart('@'))" -ForegroundColor DarkGray }
if ($useNameMatch) { Write-Host "Matching : by person name and email prefix, device's organization first" -ForegroundColor DarkGray }

# -----------------------------------------------------------------------------
# Load inventory
# -----------------------------------------------------------------------------
Write-Step "Loading NinjaOne inventory"

$devices = Get-AllDevices
Write-Ok "$($devices.Count) devices loaded"

$orgs    = Get-Organizations
$orgMap  = $orgs.ByName
$orgById = $orgs.ById
Write-Ok "$($orgMap.Count) organizations loaded"

$dir        = Get-OwnerDirectory
$owners     = $dir.ByEmail
$byIdentity = $dir.ByIdentity
Write-Ok "$($owners.Count) owner emails loaded (contacts + end users$(if ($IncludeTechnicians) { ' + technicians' }))"

# Index devices by each name field so hostname matching is forgiving but exact.
$deviceIndex = @{}
foreach ($d in $devices) {
    foreach ($field in 'systemName','displayName','dnsName','netbiosName') {
        if ($d.PSObject.Properties.Name -notcontains $field) { continue }
        $val = $d.$field
        if ([string]::IsNullOrWhiteSpace($val)) { continue }
        $key = ([string]$val).Trim().ToLowerInvariant()
        if (-not $deviceIndex.ContainsKey($key)) { $deviceIndex[$key] = [System.Collections.Generic.List[object]]::new() }
        # A device can match on several name fields; only index it once per key.
        $existingIds = @($deviceIndex[$key] | ForEach-Object { $_.id })
        if ($existingIds -notcontains $d.id) { $deviceIndex[$key].Add($d) | Out-Null }
    }
}

function Get-OrgLabel {
    param([object] $OrgId)
    if ($null -eq $OrgId) { return 'no organization' }
    if ($orgById.ContainsKey([string]$OrgId)) { return $orgById[[string]$OrgId] }
    return "org $OrgId"
}

# -----------------------------------------------------------------------------
# Plan
# -----------------------------------------------------------------------------
Write-Step "Planning"

$plan     = [System.Collections.Generic.List[object]]::new()
$seenRows = @{}   # lowercase device name -> first plan entry, for duplicate detection
$rowNum   = 1     # header is row 1 in the spreadsheet, so data starts at 2

function Add-PlanEntry {
    <# One object in both places, so flagging an earlier row as a conflict sticks. #>
    param([Parameter(Mandatory)] $Entry, [string] $SeenKey)

    $planned = [pscustomobject]$Entry
    if ($SeenKey) { $seenRows[$SeenKey] = $planned }
    $plan.Add($planned)
}

foreach ($row in $rows) {
    $rowNum++

    # Blank lines left over after pasting into the template are not rows.
    $allValues = @($row.PSObject.Properties | ForEach-Object { Clean-Value $_.Value }) -join ''
    if (-not $allValues) { continue }

    $hostName = Get-Field -Row $row -Map $colMap -Field 'Device'
    $orgName  = Get-Field -Row $row -Map $colMap -Field 'Organization'
    $email    = Get-Field -Row $row -Map $colMap -Field 'Email'
    $first    = Get-Field -Row $row -Map $colMap -Field 'FirstName'
    $last     = Get-Field -Row $row -Map $colMap -Field 'LastName'
    $fullName = Get-Field -Row $row -Map $colMap -Field 'FullName'
    $userRaw  = Get-Field -Row $row -Map $colMap -Field 'Username'
    $loginRaw = Get-Field -Row $row -Map $colMap -Field 'LastLogin'
    $osName   = Get-Field -Row $row -Map $colMap -Field 'OS'

    $sawNoUserMarker = $false
    if (Test-NoUserMarker $userRaw)  { $userRaw  = ''; $sawNoUserMarker = $true }
    if (Test-NoUserMarker $loginRaw) { $loginRaw = ''; $sawNoUserMarker = $true }
    if (Test-NoUserMarker $fullName) { $fullName = ''; $sawNoUserMarker = $true }

    $userName = Get-BareUsername $userRaw
    if (-not $userName) { $userName = Get-BareUsername $loginRaw }

    $entry = [ordered]@{
        Row           = $rowNum
        DeviceName    = $hostName
        Organization  = $orgName
        Username      = $userName
        OwnerName     = ''
        FirstName     = $first
        LastName      = $last
        Email         = $email
        EmailSource   = ''
        DeviceId      = $null
        OrgId         = $null
        OwnerUid      = $null
        OwnerType     = $null
        MatchedOn     = ''
        PreviousOwner = $null
        Action        = $null
        Reason        = $null
        Detail        = ''
    }

    if (-not $hostName) {
        $entry.Action = 'ERROR'; $entry.Reason = 'Blank device name'
        Add-PlanEntry $entry; continue
    }

    # --- Row filters --------------------------------------------------------
    if ($OnlyOrganization -and $orgName -and $orgName -ne $OnlyOrganization) {
        $entry.Action = 'FILTERED'; $entry.Reason = 'Outside -OnlyOrganization'
        Add-PlanEntry $entry; continue
    }
    if (-not $IncludeServers -and $osName -match 'server') {
        $entry.Action = 'FILTERED'; $entry.Reason = 'Server'; $entry.Detail = $osName
        Add-PlanEntry $entry; continue
    }
    if ($userName -match $script:LocalAccountPattern) {
        $entry.Action = 'FILTERED'; $entry.Reason = 'Local account, not a person'
        Add-PlanEntry $entry; continue
    }
    if ($ExcludeUserPattern) {
        $excluded = @($userName, $email, $fullName) | Where-Object { $_ -and $_ -match $ExcludeUserPattern }
        if ($excluded) {
            $entry.Action = 'FILTERED'; $entry.Reason = 'Matched -ExcludeUserPattern'
            Add-PlanEntry $entry; continue
        }
    }

    # --- Derive the person's name ------------------------------------------
    if (-not $first -or -not $last) {
        $split = Split-PersonName $fullName
        if ($null -eq $split -and $userName) {
            # No display name at all: fall back to the account name, splitting
            # NADSA.SOLCAct1 / bw6235 style values on the dot.
            if ($userName -match '\.') {
                $split = Split-PersonName (($userName -split '\.') -join ' ')
            } else {
                $split = Split-PersonName $userName
            }
        }
        if ($null -ne $split) {
            if (-not $first) { $first = $split.First }
            if (-not $last)  { $last  = $split.Last  }
        }
    }
    $entry.FirstName = $first
    $entry.LastName  = $last

    # --- Resolve the email -------------------------------------------------
    if ($email) {
        $entry.EmailSource = 'csv'
    } elseif ($EmailDomain) {
        $email = New-DerivedEmail -First $first -Last $last -Username $userName `
                                  -Domain $EmailDomain -Pattern $EmailPattern
        if ($email) { $entry.EmailSource = "derived ($EmailPattern)" }
    }
    $entry.Email = $email

    $hasPerson = $email -or $first -or $last -or $userName
    if (-not $hasPerson -and $sawNoUserMarker) {
        $entry.Action = 'NOUSER'; $entry.Reason = 'Nobody logged in when the sheet was exported'
        Add-PlanEntry $entry; continue
    }

    # A row with no person data at all is a spreadsheet gap (usually a failed
    # VLOOKUP that left #VALUE! behind), not something to guess at.
    if (-not $hasPerson) {
        $entry.Action = 'ERROR'
        $entry.Reason = 'No user on this row (blank or an Excel error)'
        Add-PlanEntry $entry; continue
    }

    if (-not $email -and -not $useNameMatch) {
        $entry.Action = 'ERROR'
        $entry.Reason = 'No usable email'
        $entry.Detail = 'Add an Email column, pass -EmailDomain to derive one, or use -MatchByName.'
        Add-PlanEntry $entry; continue
    }

    # --- Resolve device to exactly one match -------------------------------
    # NOTE: do not name this '$matches' - that is a PowerShell automatic
    # variable that every -match operation overwrites.
    # Assigned in two statements on purpose: '$x = if (...) { @() }' collapses an
    # empty array to $null in PowerShell, which then fails on .Count.
    $key           = $hostName.ToLowerInvariant()
    $deviceMatches = @()
    if ($deviceIndex.ContainsKey($key)) { $deviceMatches = @($deviceIndex[$key]) }

    if ($deviceMatches.Count -eq 0) {
        $entry.Action = 'ERROR'; $entry.Reason = 'Device not found in NinjaOne'
        Add-PlanEntry $entry; continue
    }

    # Narrow duplicate hostnames using the Organization column when we have one.
    if ($deviceMatches.Count -gt 1 -and $orgName -and $orgMap.ContainsKey($orgName.ToLowerInvariant())) {
        $wantOrg = $orgMap[$orgName.ToLowerInvariant()]
        $narrowed = @($deviceMatches | Where-Object { $_.organizationId -eq $wantOrg })
        if ($narrowed.Count -eq 1) { $deviceMatches = $narrowed }
    }

    if ($deviceMatches.Count -gt 1) {
        $entry.Action = 'ERROR'
        $entry.Reason = 'More than one device has this name'
        $entry.Detail = "Device ids: $(($deviceMatches | ForEach-Object { $_.id }) -join ', ')"
        Add-PlanEntry $entry; continue
    }

    $device              = $deviceMatches[0]
    $entry.DeviceId      = $device.id
    $entry.OrgId         = $device.organizationId
    $currentOwner        = if ($device.PSObject.Properties.Name -contains 'assignedOwnerUid') { $device.assignedOwnerUid } else { $null }
    $entry.PreviousOwner = $currentOwner

    # --- Duplicate device rows ---------------------------------------------
    # Two rows for one device would silently fight over the assignment.
    $personKey = if ($email) { $email.ToLowerInvariant() } elseif ($userName) { Get-IdentityKey $userName } else { Get-IdentityKey "$first $last" }
    $personTag = if ($email) { $email } elseif ($userName) { $userName } else { "$first $last" }
    $entry['PersonKey'] = $personKey

    if ($seenRows.ContainsKey($key)) {
        $prev = $seenRows[$key]
        if ($prev.PersonKey -eq $personKey) {
            $entry.Action = 'SKIP'; $entry.Reason = "Duplicate row, same device and user"; $entry.Detail = "Same as row $($prev.Row)"
        } else {
            $entry.Action = 'ERROR'
            $entry.Reason = 'Same device listed with two different users'
            $entry.Detail = "Row $($prev.Row) says '$($prev.PersonTag)', this row says '$personTag'"
            $prev.Action  = 'ERROR'
            $prev.Reason  = 'Same device listed with two different users'
            $prev.Detail  = "This row says '$($prev.PersonTag)', row $rowNum says '$personTag'"
        }
        $entry['PersonTag'] = $personTag
        Add-PlanEntry $entry; continue
    }
    $entry['PersonTag'] = $personTag

    # --- Resolve owner ------------------------------------------------------
    $owner = $null
    if ($email -and $owners.ContainsKey($email.ToLowerInvariant())) {
        $owner = $owners[$email.ToLowerInvariant()]
        $entry.MatchedOn = 'email'
    }

    if (-not $owner -and $useNameMatch) {
        $lookupKeys = @((Get-IdentityKey $userName), (Get-IdentityKey "$first $last")) |
                      Where-Object { $_ } | Select-Object -Unique
        $candidates = @()
        foreach ($lookup in $lookupKeys) {
            if ($byIdentity.ContainsKey($lookup)) { $candidates += @($byIdentity[$lookup]) }
        }

        if ($candidates.Count -gt 0) {
            $pick = Select-OwnerForDevice -Candidates $candidates -DeviceOrgId $device.organizationId

            if ($pick.Problem -eq 'AMBIGUOUS') {
                $entry.Action = 'ERROR'
                $entry.Reason = 'More than one user matches this name'
                $entry.Detail = (@($pick.Others | ForEach-Object { "$($_.FirstName) $($_.LastName) <$($_.Email)>" }) -join '; ')
                Add-PlanEntry $entry $key; continue
            }
            if ($pick.Problem -eq 'OTHER_ORG') {
                $entry.Action = 'ERROR'
                $entry.Reason = 'User exists, but in a different organization than the device'
                $entry.Detail = "Device is in '$(Get-OrgLabel $device.organizationId)'; user: " +
                                (@($pick.Others | ForEach-Object { "$($_.FirstName) $($_.LastName) in '$(Get-OrgLabel $_.OrganizationId)'" }) -join '; ')
                Add-PlanEntry $entry $key; continue
            }

            $owner = $pick.Owner
            $entry.MatchedOn = if ((Get-IdentityKey "$($owner.FirstName) $($owner.LastName)") -in $lookupKeys) { 'name' } else { 'email prefix' }
            # Report the owner's real address, not a derived guess that missed.
            $entry.Email       = $owner.Email
            $entry.EmailSource = 'from matched owner'
            $email             = $owner.Email
        }
    }

    if (-not $owner) {
        if (-not $CreateMissingContacts) {
            $entry.Action = 'ERROR'
            $entry.Reason = 'No matching user in NinjaOne'
            $entry.Detail = "Looked for '$personTag' in contacts and end users"
            Add-PlanEntry $entry $key; continue
        }
        if (-not $email) {
            $entry.Action = 'ERROR'
            $entry.Reason = 'Contact must be created but there is no email to create it with'
            $entry.Detail = 'Pass -EmailDomain, or add the user in NinjaOne.'
            Add-PlanEntry $entry $key; continue
        }
        if (-not $first -or -not $last) {
            $entry.Action = 'ERROR'
            $entry.Reason = 'Contact must be created but the name could not be determined'
            Add-PlanEntry $entry $key; continue
        }
        $entry.Action    = 'CREATE+ASSIGN'
        $entry.OwnerType = 'Contact (new)'
        $entry.MatchedOn = 'new'
        $entry.OwnerName = "$first $last"
        $entry.Reason    = 'New contact, then assign'
        $entry.Detail    = "Will create the contact in org $($device.organizationId)"
        Add-PlanEntry $entry $key; continue
    }

    $entry.OwnerUid  = $owner.Uid
    $entry.OwnerType = $owner.Type
    $entry.OwnerName = "$($owner.FirstName) $($owner.LastName)"

    # --- Organization validation -------------------------------------------
    # Owner org must be null (global) or exactly the device's org.
    if ($null -ne $owner.OrganizationId -and $owner.OrganizationId -ne $device.organizationId) {
        $entry.Action = 'ERROR'
        $entry.Reason = 'User exists, but in a different organization than the device'
        $entry.Detail = "User is in '$(Get-OrgLabel $owner.OrganizationId)', device is in '$(Get-OrgLabel $device.organizationId)'"
        Add-PlanEntry $entry $key; continue
    }

    # --- Already correct? ---------------------------------------------------
    if ($currentOwner -and $currentOwner -eq $owner.Uid) {
        $entry.Action = 'SKIP'; $entry.Reason = 'Already assigned to this user'
        Add-PlanEntry $entry $key; continue
    }

    $entry.Action = if ($currentOwner) { 'REASSIGN' } else { 'ASSIGN' }
    $entry.Reason = if ($currentOwner) { 'Replacing existing assignment' } else { 'Currently unassigned' }
    Add-PlanEntry $entry $key
}

# -----------------------------------------------------------------------------
# Dry-run report
# -----------------------------------------------------------------------------
Write-Step "Dry-run report"

$reportColumns = 'Row','DeviceName','Organization','Username','OwnerName','Email','MatchedOn','DeviceId','OrgId',
                 'OwnerUid','OwnerType','PreviousOwner','Action','Reason','Detail'

if ($plan.Count -le $script:MaxRowsToPrint) {
    $plan | Format-Table Row, DeviceName, Username, OwnerName, MatchedOn, Action, Reason -AutoSize |
        Out-String -Width 4096 | Write-Host
}

$toWrite    = @($plan | Where-Object { $_.Action -in 'ASSIGN','REASSIGN','CREATE+ASSIGN' })
$toSkip     = @($plan | Where-Object { $_.Action -eq 'SKIP' })
$toFiltered = @($plan | Where-Object { $_.Action -eq 'FILTERED' })
$toNoUser   = @($plan | Where-Object { $_.Action -eq 'NOUSER' })
$toError    = @($plan | Where-Object { $_.Action -eq 'ERROR' })

Write-Host "Planned changes : $($toWrite.Count)" -ForegroundColor Yellow
Write-Host "Already correct : $($toSkip.Count)"  -ForegroundColor DarkGray
Write-Host "No user on row  : $($toNoUser.Count)" -ForegroundColor DarkGray
Write-Host "Filtered out    : $($toFiltered.Count)" -ForegroundColor DarkGray
if ($toFiltered.Count) {
    $toFiltered | Group-Object Reason | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("  {0,5} x {1}" -f $_.Count, $_.Name) -ForegroundColor DarkGray
    }
}
Write-Host "Rows in error   : $($toError.Count)" -ForegroundColor $(if ($toError.Count) { 'Red' } else { 'DarkGray' })

if ($toError.Count) {
    Write-Host "`nError breakdown:" -ForegroundColor Red
    $toError | Group-Object Reason | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("  {0,5} x {1}" -f $_.Count, $_.Name)
    }

    Write-Host "`nError rows:" -ForegroundColor Red
    $toError | Select-Object -First 25 | ForEach-Object {
        Write-Host ("  Row {0}: {1} / {2} - {3}" -f $_.Row, $_.DeviceName, $_.Username, $(if ($_.Detail) { $_.Detail } else { $_.Reason }))
    }
    if ($toError.Count -gt 25) { Write-Host "  ... and $($toError.Count - 25) more in the report." }
}

if (-not $Execute) {
    $plan | Select-Object $reportColumns | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nDRY RUN complete. No changes were made." -ForegroundColor Green
    Write-Host "Report written to: $LogPath"
    Write-Host "Re-run with -Execute to apply the $($toWrite.Count) planned change(s)."
    return
}

if ($toWrite.Count -eq 0) {
    Write-Host "`nNothing to do - no rows require a change." -ForegroundColor Green
    $plan | Select-Object $reportColumns | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "Report written to: $LogPath"
    return
}

# -----------------------------------------------------------------------------
# Confirmation gate. An automation has no console to answer a prompt, so there
# the -Execute parameter is the confirmation. Secure custom fields can only be read
# inside an automation, so getting credentials from them proves that is where we are.
# -----------------------------------------------------------------------------
if (-not $Force -and -not $script:Credentials.FromFields -and [Environment]::UserInteractive) {
    Write-Host ""
    $answer = ''
    try {
        $answer = Read-Host "Apply $($toWrite.Count) assignment change(s) to $script:BaseUrl? Type YES to proceed"
    } catch {
        Write-Host "Could not prompt for confirmation in this session. Add -Force to run unattended." -ForegroundColor Yellow
    }
    if ($answer -cne 'YES') {
        Write-Host "Aborted by operator. No changes made." -ForegroundColor Yellow
        return
    }
}

# -----------------------------------------------------------------------------
# Execute
# -----------------------------------------------------------------------------
Write-Step "Applying changes"

$printRows = $plan.Count -le $script:MaxRowsToPrint
$ok = 0; $failed = 0

function Add-Result {
    param($Item, [string]$Action, [string]$Result, $NewOwnerUid, [string]$Detail)
    $script:Results.Add([pscustomobject]@{
        Row = $Item.Row; DeviceName = $Item.DeviceName; DeviceId = $Item.DeviceId; OrgId = $Item.OrgId
        Username = $Item.Username; OwnerName = $Item.OwnerName; Email = $Item.Email
        Action = $Action; Result = $Result
        PreviousOwnerUid = $Item.PreviousOwner; NewOwnerUid = $NewOwnerUid
        Detail = $Detail
    })
}

foreach ($item in $plan) {

    if ($item.Action -in 'SKIP','FILTERED','NOUSER') {
        if ($printRows) { Write-Skip "$($item.DeviceName) - $($item.Reason)" }
        Add-Result $item $item.Action 'SKIPPED' $item.PreviousOwner $item.Reason
        continue
    }

    if ($item.Action -eq 'ERROR') {
        # Already listed above. A row that was never attempted is not a failed write.
        $detail = if ($item.Detail) { "$($item.Reason): $($item.Detail)" } else { $item.Reason }
        Add-Result $item 'ERROR' 'NOT ATTEMPTED' $null $detail
        continue
    }

    $ownerUid = $item.OwnerUid

    # Create the contact first when needed.
    if ($item.Action -eq 'CREATE+ASSIGN') {
        $created = New-NinjaContact -FirstName $item.FirstName -LastName $item.LastName `
                                    -Email $item.Email -OrganizationId $item.OrgId

        if ($created.AlreadyExists) {
            # Email exists elsewhere in the tenant - almost always an org mismatch.
            Write-Err "$($item.DeviceName) - contact email already in use in another organization; resolve manually"
            $failed++
            Add-Result $item 'CREATE' 'FAILED' $null 'EMAIL_ALREADY_IN_USE (409) - existing record is in a different org'
            continue
        }
        if (-not $created.Success) {
            Write-Err "$($item.DeviceName) - contact creation failed: $($created.Error)"
            $failed++
            Add-Result $item 'CREATE' 'FAILED' $null $created.Error
            continue
        }

        $ownerUid = $created.Data.uid
        if ($printRows) { Write-Ok "$($item.Email) - contact created (id $($created.Data.id), org $($item.OrgId))" }
    }

    # Assign.
    $res = Invoke-NinjaApi -Method POST -Path "/device/$($item.DeviceId)/owner/$ownerUid"

    if ($res.Success) {
        $ok++
        if ($printRows) { Write-Ok "$($item.DeviceName) (id $($item.DeviceId)) -> $($item.OwnerName)" }
        Add-Result $item $item.Action 'SUCCESS' $ownerUid ''
    }
    else {
        $failed++
        $hint = ''
        if ($res.Error -match 'owner_has_not_same_node_organization') {
            $hint = ' (owner org does not match device org)'
        }
        Write-Err "$($item.DeviceName) (id $($item.DeviceId)) - HTTP $($res.Status)$hint"
        Add-Result $item $item.Action 'FAILED' $null "HTTP $($res.Status): $($res.Error)$hint"
    }
}

# -----------------------------------------------------------------------------
# Summary
# -----------------------------------------------------------------------------
$script:Results | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8

Write-Step "Summary"
Write-Host "  Assigned       : $ok"              -ForegroundColor Green
Write-Host "  Skipped        : $($toSkip.Count + $toFiltered.Count + $toNoUser.Count)" -ForegroundColor DarkGray
Write-Host "  Rows in error  : $($toError.Count)" -ForegroundColor $(if ($toError.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host "  Failed writes  : $failed"          -ForegroundColor $(if ($failed) { 'Red' } else { 'DarkGray' })
Write-Host "  Log            : $LogPath"

# Exit 1 marks the automation failed in NinjaOne, so keep it for writes the API
# refused. Sheet rows that need fixing are in the log and do not fail the run.
if ($failed -gt 0) {
    Write-Host "`nTo roll back a device, clear its assignment with:" -ForegroundColor Yellow
    Write-Host '  DELETE /v2/device/{deviceId}/owner'
    Write-Host "The PreviousOwnerUid column in the log holds the value to restore."
    exit 1
}
