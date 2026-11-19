# NinjaOne Bulk Assigned User

Sets the **Assigned User** (device owner) on NinjaOne devices from a CSV. Each row is matched to one device and one existing contact, end user or technician, and the owner is set with `POST /v2/device/{id}/owner/{ownerUid}`.

The script starts in dry-run mode. It writes a report of what it would change and does nothing else until you pass `-Execute`.

## Files

| File | What it is |
|---|---|
| `Set-NinjaAssignedUser.ps1` | The bulk script |
| `Set-NinjaAssignedUser-Minimal.ps1` | About 40 lines that assign one owner to one device. Use it to see the API calls, or to test your credentials |
| `samples/sample-assigned-users.csv` | Full example with org, OS, name and email columns |
| `samples/sample-assigned-users-minimal.csv` | Smallest useful sheet: device, last login, username, name |

## Requirements

- Windows PowerShell 5.1 or PowerShell 7
- A NinjaOne API client (step 1 below)
- The owners must already exist in NinjaOne as contacts, end users or technicians. `-CreateMissingContacts` can create contacts for you

## Setup

### 1. API client

**Administration → Apps → API → Client app IDs → Add.**

| Field | Value |
|---|---|
| Application Platform | **API Services (Machine to Machine)** |
| Scopes | **Monitoring** and **Management** |
| Grant type | **Client Credentials** |

The secret is only shown once, so copy it now.

### 2. Choose how to run it

**Option A: as a NinjaOne automation (no credentials on disk)**

1. Create two **device** custom fields, type **Secure**, Automations **Read Only**:

   | Field name | Value |
   |---|---|
   | `ninjaoneClientId` | Client ID |
   | `ninjaoneClientSecret` | Client secret |

2. Fill both in on the device that will run the script.
3. Save your CSV on that device as `C:\Temp\Assigned-Users.csv`.
4. **Administration → Library → Automation → Add → New Script**: PowerShell, Windows, Run As **System**, paste `Set-NinjaAssignedUser.ps1`.
5. Run it on that device. Put the parameters in the **Parameters** box, for example `-Instance eu.ninjaone.com` for a dry run, then `-Instance eu.ninjaone.com -Execute` to apply.

**Option B: from a PowerShell window**

```powershell
$env:NINJA_CLIENT_ID     = '<client id>'
$env:NINJA_CLIENT_SECRET = '<client secret>'

# Dry run
.\Set-NinjaAssignedUser.ps1 -CsvPath .\Assigned-Users.csv -Instance app.ninjarmm.com

# Apply
.\Set-NinjaAssignedUser.ps1 -CsvPath .\Assigned-Users.csv -Instance app.ninjarmm.com -Execute
```

### 3. Set your instance

`-Instance` is the host you sign in to: `app.ninjarmm.com` (US, the default), `us2.ninjarmm.com`, `eu.ninjarmm.com`, `ca.ninjarmm.com`, `oc.ninjarmm.com`, and so on. A full URL also works.

## The CSV

Headers are matched without regard to case, spaces or underscores. Only the device column is required, plus at least one column that identifies the person.

| Meaning | Accepted headers |
|---|---|
| Device (required) | `Agent Name`, `Device`, `DeviceName`, `Hostname`, `SystemName` |
| Organization | `Client Name`, `Organization`, `Org`, `Client` |
| Username | `Agent User`, `Username`, `SamAccountName` |
| Last login | `LastLogin`, `LastLoggedOnUser`, `LoggedOnUser` |
| OS | `Agent Operating System`, `OperatingSystem`, `OS` |
| Email | `Email`, `EmailAddress`, `UserEmail`, `Mail` |
| First / last name | `FirstName`, `GivenName` / `LastName`, `Surname`, `FamilyName` |
| Full name | `Name`, `FullName`, `DisplayName` |

Example:

```csv
Device,Organization,Last Login,Username,Name,Email
PCW117,North Region,AP\mw6235,mw6235,Marcus Webb,Marcus.Webb@example.com
LT705,EMEA,NA\pr1227,pr1227,Priya Raman,Priya.Raman@example.com
```

The headers in a ConnectWise Automate agent export (`Agent Name`, `Client Name`, `Agent User`) work without any changes.

Save the sheet as **CSV UTF-8**. An `.xlsx` renamed to `.csv` is detected and rejected with a message.

## How owners are matched

1. **Email.** A row with an email matches the owner with that exact address.
2. **Name.** With no email, the row matches on the person's name, or on the part of the owner's email before the `@`. Case, dots and accents are ignored. `DOMAIN\First.Last` in a login column is read as *First Last*.
3. **Ties.** When the same name exists more than once, the owner in the device's organization wins, then an owner with no organization. Anything still tied is reported as an error. The script does not guess.

NinjaOne only accepts an owner who has no organization or who is in the **same organization as the device**. Any other pairing is reported before anything is written (the API would return `403 owner_has_not_same_node_organization`).

## What the report says

Every row gets one action in the report:

| Action | Meaning |
|---|---|
| `ASSIGN` | Device has no owner and will get one |
| `REASSIGN` | Device has a different owner, which will be replaced |
| `SKIP` | Already assigned to this owner |
| `NOUSER` | Row has no user (blank, `Not Logged In`). Device left alone |
| `FILTERED` | Local or service account (`admin`, `administrator`, `guest`, …), or a Windows Server without `-IncludeServers` |
| `ERROR` | Device not found, owner not found, duplicate match, or wrong organization. The `Reason` column says which |

The report is saved next to the CSV as `ninja-assigned-user-<timestamp>.csv`. After `-Execute`, the log includes `PreviousOwnerUid` for every change.

## Parameters

| Parameter | Default | What it does |
|---|---|---|
| `-CsvPath` | `C:\Temp\Assigned-Users.csv` | Input file |
| `-Instance` | `app.ninjarmm.com` | NinjaOne host or URL |
| `-Execute` | off | Make the changes. Without it, dry run only |
| `-OnlyOrganization` | | Only process rows for one organization. Good for a pilot |
| `-IncludeServers` | off | Also process rows whose OS says Server |
| `-IncludeTechnicians` | off | Also match technicians as owners |
| `-MatchByName` | off | When the sheet has emails, fall back to the name for rows whose email is blank or unknown |
| `-EmailDomain` | | Build missing emails from the name, e.g. `-EmailDomain example.com` |
| `-EmailPattern` | `first.last` | `first.last`, `firstlast`, `flast`, `first_last` or `username` |
| `-LowercaseEmail` | off | Lowercase the derived emails |
| `-CreateMissingContacts` | off | Create a contact in the device's organization when no owner matches. Needs an email |
| `-ExcludeUserPattern` | | Extra regex of accounts to skip, e.g. `'adm$\|^svc'` |
| `-ClientIdField` / `-ClientSecretField` | `ninjaoneClientId` / `ninjaoneClientSecret` | Custom field names for the credentials |
| `-LogPath` | next to the CSV | Where to write the report |
| `-Force` | off | Skip the confirmation prompt in an interactive window |

## Undo

Each change in the log has the device ID and `PreviousOwnerUid`. To restore an owner, call `POST /v2/device/{deviceId}/owner/{PreviousOwnerUid}`. To clear one, call `DELETE /v2/device/{deviceId}/owner`.

## Notes

- Credentials come from the secure custom fields when the script runs as an automation, otherwise from `NINJA_CLIENT_ID` / `NINJA_CLIENT_SECRET`. Tokens and secrets never go into the console output or the log.
- A wrong client ID makes `/ws/oauth/token` return **404**, not 401. Check the ID and the instance before looking for a URL problem.
- API rate limits (429) are retried automatically.
- The file is read with shared access, so it can stay open in Excel or syncing in OneDrive.

## License

MIT
