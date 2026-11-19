# Minimal working example: assign one existing contact to one device by hostname.
# Prereqs: $env:NINJA_CLIENT_ID / $env:NINJA_CLIENT_SECRET, API app with
#          Client Credentials grant and scopes "monitoring management".

$Instance   = 'app.ninjaone.com'      # your region, e.g. eu.ninjaone.com / ca.ninjarmm.com
$HostName   = 'PC-001'
$OwnerEmail = 'john.smith@example.com'

# 1. Token
$token = (Invoke-RestMethod -Method Post -Uri "https://$Instance/ws/oauth/token" `
    -ContentType 'application/x-www-form-urlencoded' -Body @{
        grant_type    = 'client_credentials'
        client_id     = $env:NINJA_CLIENT_ID
        client_secret = $env:NINJA_CLIENT_SECRET
        scope         = 'monitoring management'
    }).access_token

$headers = @{ Authorization = "Bearer $token" }

# 2. Find the device by hostname (exact, case-insensitive)
$device = Invoke-RestMethod -Uri "https://$Instance/v2/devices?pageSize=1000" -Headers $headers |
          Where-Object { $_.systemName -eq $HostName }

if (-not $device)            { throw "No device named '$HostName'" }
if ($device.Count -gt 1)     { throw "'$HostName' matched $($device.Count) devices - resolve manually" }

# 3. Find the contact by email
$contact = Invoke-RestMethod -Uri "https://$Instance/v2/contacts" -Headers $headers |
           Where-Object { $_.email -eq $OwnerEmail }

if (-not $contact) { throw "No contact with email '$OwnerEmail'" }

# 4. Org guard - owner org must be absent (global) or equal the device's org
if ($null -ne $contact.organizationId -and $contact.organizationId -ne $device.organizationId) {
    throw "Org mismatch: contact is in org $($contact.organizationId), device is in org $($device.organizationId)"
}

# 5. Assign
Invoke-RestMethod -Method Post -Headers $headers `
    -Uri "https://$Instance/v2/device/$($device.id)/owner/$($contact.uid)"
# -> @{ success = True }

# 6. Verify
$after = Invoke-RestMethod -Uri "https://$Instance/v2/device/$($device.id)" -Headers $headers
"assignedOwnerUid = $($after.assignedOwnerUid)"      # should equal $contact.uid

# To clear the assignment:
# Invoke-RestMethod -Method Delete -Headers $headers -Uri "https://$Instance/v2/device/$($device.id)/owner"
