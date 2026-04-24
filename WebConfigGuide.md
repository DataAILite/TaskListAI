# TaskListAI web.config Configuration Guide

This guide explains how to configure the `web.config` file for your TaskListAI (HelpDesk) installation. Every setting you need to change is marked with `[placeholder text]` in the default file.

**Before you start**, make a backup:

```
copy web.config web.config.backup
```

---

## Step 1: Connection Strings

The `<connectionStrings>` section defines how TaskListAI connects to your database. The file ships with MySQL active by default and all other database templates commented out inside `<!-- ... -->`.

There are two connection names:

1. **MySqlConnection** — TaskListAI operational database (stores tickets, users, configuration)
2. **UserSqlConnection** — your data database (the data your helpdesk works with)

To switch database engines, comment out the active MySQL lines and uncomment the block for your engine. Replace all `[placeholder]` values with your actual server details.

### MySQL (active by default)

```xml
<add name="MySqlConnection"
     connectionString="Server=YOUR_SERVER; Database=YOUR_OPERATIONAL_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD;"
     providerName="MySql.Data.MySqlClient" />
<add name="UserSqlConnection"
     connectionString="Server=YOUR_DATA_SERVER; Database=YOUR_DATA_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD;"
     providerName="MySql.Data.MySqlClient" />
```

### SQL Server

```xml
<add name="MySqlConnection"
     connectionString="Server=YOUR_SERVER; Database=YOUR_OPERATIONAL_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD; Trusted_Connection=True"
     providerName="System.Data.SqlClient" />
<add name="UserSqlConnection"
     connectionString="Server=YOUR_DATA_SERVER; Database=YOUR_DATA_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD; Trusted_Connection=True"
     providerName="System.Data.SqlClient" />
```

### PostgreSQL

Default port is 5432.

```xml
<add name="MySqlConnection"
     connectionString="Server=YOUR_SERVER; Port=5432; Database=YOUR_OPERATIONAL_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD;"
     providerName="Npgsql" />
<add name="UserSqlConnection"
     connectionString="Server=YOUR_DATA_SERVER; Port=5432; Database=YOUR_DATA_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD;"
     providerName="Npgsql" />
```

### Oracle

Oracle requires an additional `SystemSqlConnection` for system-level access with SYSDBA privilege for system operations.

```xml
<add name="SystemSqlConnection"
     connectionString="Data Source=YOUR_DATA_SOURCE; User ID=SYS; Password=YOUR_SYS_PASSWORD; DBA Privilege=SYSDBA"
     providerName="Oracle.ManagedData.Client" />
<add name="MySqlConnection"
     connectionString="Data Source=YOUR_DATA_SOURCE; User ID=YOUR_USER; Password=YOUR_PASSWORD;"
     providerName="Oracle.ManagedData.Client" />
<add name="UserSqlConnection"
     connectionString="Data Source=YOUR_DATA_SOURCE; User ID=YOUR_USER; Password=YOUR_PASSWORD;"
     providerName="Oracle.ManagedData.Client" />
```

**License required:** Oracle Database requires a valid commercial license from Oracle Corporation. TaskListAI does not provide or distribute Oracle software. Visit https://www.oracle.com for licensing information.

### InterSystems IRIS

IRIS requires an additional `SystemSqlConnection` for system-level access. Default ports are 1972 (system) and 51773 (data).

```xml
<add name="SystemSqlConnection"
     connectionString="Server=YOUR_SERVER; Port=1972; Namespace=%SYS; User ID=_SYSTEM; Password=YOUR_SYS_PASSWORD"
     providerName="InterSystems.Data.IRISClient" />
<add name="MySqlConnection"
     connectionString="Server=YOUR_SERVER; Port=51773; Namespace=YOUR_NAMESPACE; User ID=YOUR_USER; Password=YOUR_PASSWORD"
     providerName="InterSystems.Data.IRISClient" />
<add name="UserSqlConnection"
     connectionString="Server=YOUR_DATA_SERVER; Port=51773; Namespace=YOUR_NAMESPACE; User ID=YOUR_USER; Password=YOUR_PASSWORD"
     providerName="InterSystems.Data.IRISClient" />
```

**License required:** InterSystems IRIS requires a valid commercial license from InterSystems Corporation. TaskListAI does not provide or distribute InterSystems software. Visit https://www.intersystems.com for licensing information.

### InterSystems Caché

Same format as IRIS, but use `providerName="InterSystems.Data.CacheClient"` and port 1972. No `SystemSqlConnection` is needed.

```xml
<add name="MySqlConnection"
     connectionString="Server=YOUR_SERVER; Port=1972; Namespace=YOUR_NAMESPACE; User ID=YOUR_USER; Password=YOUR_PASSWORD"
     providerName="InterSystems.Data.CacheClient" />
<add name="UserSqlConnection"
     connectionString="Server=YOUR_DATA_SERVER; Port=1972; Namespace=YOUR_NAMESPACE; User ID=YOUR_USER; Password=YOUR_PASSWORD"
     providerName="InterSystems.Data.CacheClient" />
```

**License required:** InterSystems Caché requires a valid commercial license from InterSystems Corporation.

---

## Step 2: Application Identity

In the `<appSettings>` section, set your application name and unit:

```xml
<add key="ourapplication" value="HelpDesk" />
<add key="unit" value="MyCompany" />
<add key="pagettl" value="HelpDesk with Analytical Intelligence" />
```

- `ourapplication` — your application display name (e.g., `HelpDesk`)
- `unit` — your organization or unit name
- `pagettl` — the title shown in the browser tab

---

## Step 3: Super User Password

```xml
<add key="superpass" value="YourSecurePassword" />
```

This is the password for the built-in super administrator account. Choose a strong password and keep it confidential.

---

## Step 4: Web URLs

Set all URL keys to the full address where your TaskListAI site is hosted:

```xml
<add key="unitOUReportsWeb" value="https://yourserver/TaskListAI/" />
<add key="unitRegistrationWeb" value="https://yourserver/TaskListAI/" />
<add key="webour" value="https://yourserver/TaskListAI/" />
<add key="weboureports" value="https://yourserver/TaskListAI/" />
<add key="webhelpdesk" value="https://yourserver/TaskListAI/" />
```

---

## Step 5: Database Server References

```xml
<add key="OUReportsServer" value="your-db-server-hostname" />
<add key="unitOURdbConnStr" value="Server=YOUR_SERVER; Database=YOUR_DB; User ID=YOUR_USER; Password=YOUR_PASSWORD;" />
```

- `OUReportsServer` — hostname or IP of your operational database server
- `unitOURdbConnStr` — full connection string for unit-level database access (should match your `MySqlConnection` credentials)

---

## Step 6: Email (SMTP)

Configure email in two places. Both must use the same email address and password.

**In `<appSettings>`:**

```xml
<add key="SmtpCred" value="smtp.gmail.com" />
<add key="smtpemail" value="noreply@yourcompany.com" />
<add key="smtpemailpass" value="your-app-password" />
<add key="supportemail" value="support@yourcompany.com" />
```

**In `<system.net>` / `<mailSettings>`:**

```xml
<smtp deliveryMethod="Network" from="noreply@yourcompany.com">
  <network defaultCredentials="false"
           host="smtp.gmail.com"
           password="your-app-password"
           port="587"
           userName="noreply@yourcompany.com"
           enableSsl="true" />
</smtp>
```

**Gmail users:** You must generate an App Password at https://myaccount.google.com/apppasswords (requires 2-Step Verification enabled). Your regular Gmail password will not work.

---

## Step 7: File Upload Folder

```xml
<add key="fileupload" value="C:\TaskListAI\uploads\" />
```

Set this to a folder on the server where TaskListAI can store uploaded files. The IIS application pool identity must have read and write permissions on this folder.

---

## Step 8: Google Maps API Key (optional)

Only needed if your reports use geographic map visualizations.

```xml
<add key="mapkey" value="AIzaSy..." />
```

Get a key at https://console.cloud.google.com/google/maps-api. If you are not using maps, leave the placeholder value.

---

## Step 9: OpenAI Integration (optional)

Only needed if you want AI-powered data analysis and chat features.

```xml
<add key="openaikey" value="sk-proj-..." />
<add key="openaiorganization" value="org-..." />
<add key="apiURL" value="https://api.openai.com/v1/chat/completions" />
<add key="openaimodel" value="gpt-4o" />
<add key="openaimaxTokens" value="128000" />
```

Available models: `gpt-4o`, `gpt-4o-mini`, `o3`, `o3-mini`. Get your API key at https://platform.openai.com/api-keys.

If you are not using AI features, leave the placeholder values.

---

## Step 10: Database Provider Flags

Set the provider flag to `OK` for each database engine available on your server. Leave the others empty.

```xml
<add key="MySqlProv" value="OK" />       <!-- MySQL -->
<add key="SQLServerProv" value="OK" />   <!-- SQL Server -->
<add key="CacheProv" value="" />         <!-- InterSystems Caché -->
<add key="IRISProv" value="" />          <!-- InterSystems IRIS -->
<add key="CSVProv" value="OK" />         <!-- CSV flat file import -->
<add key="Oracle" value="" />            <!-- Oracle -->
<add key="ODBC" value="OK" />            <!-- ODBC generic -->
<add key="OleDb" value="OK" />           <!-- OLE DB generic -->
<add key="Npgsql" value="OK" />          <!-- PostgreSQL -->
```

The default file ships with MySQL, SQL Server, CSV, ODBC, OLE DB, and PostgreSQL enabled.

---

## Step 11: Database Case Sensitivity

These control how TaskListAI formats table and column names in SQL queries.

```xml
<add key="ourdbcase" value="lower" />   <!-- operational database -->
<add key="csvdbcase" value="lower" />   <!-- CSV import database -->
<add key="userdbcase" value="" />       <!-- user data database -->
```

Recommended values by database engine:

| Database | Recommended Setting |
|----------|-------------------|
| MySQL | `lower` |
| PostgreSQL | `lower` |
| SQL Server | `mix` |
| Oracle | `upper` |
| InterSystems IRIS / Caché | `upper` |

---

## Settings You Usually Don't Need to Change

| Key | Default | What It Does |
|-----|---------|--------------|
| `webinstall` | `OURweb` | Internal install marker |
| `dbinstall` | `OURdb` | Internal install marker |
| `unitenddate` | `2040-12-31 23:59:00` | License expiration |
| `UnitAuthenticate` | `NO` | Unit-level authentication toggle |
| `maxretries` | `5` | API call retry attempts |
| `SiteFor` | `Production` | Environment label |
| `DaysFree` | `2000` | Free trial days for new users |
| `version` | `33-00` | Application version (do not change) |
| `dataaidownpay` | `$10` | Download fee (if payments enabled) |
| `dataaidbpay` | `$100` | Database setup fee |
| `ACEOLEDBversion` | `Provider=Microsoft.ACE.OLEDB.16.0;` | ACE driver version for Excel/Access imports |

---

## Quick Checklist

After editing, verify you have completed each step:

- [ ] Set your database connection strings with real credentials (`MySqlConnection` and `UserSqlConnection`)
- [ ] Set `ourapplication` to your app name
- [ ] Set `superpass` to a strong password
- [ ] Set `pagettl` to your site title
- [ ] Set all five web URL keys (`unitOUReportsWeb`, `unitRegistrationWeb`, `webour`, `weboureports`, `webhelpdesk`)
- [ ] Set `OUReportsServer` and `unitOURdbConnStr`
- [ ] Configured SMTP email in both `<appSettings>` and `<system.net>`
- [ ] Set `fileupload` to a writable folder path
- [ ] Set `mapkey` (if using map reports)
- [ ] Set `openaikey`, `openaimodel`, and `apiURL` (if using AI features)
- [ ] Enabled the correct database provider flags
- [ ] Set `ourdbcase`, `csvdbcase`, `userdbcase` for your database engine
- [ ] Saved the file and restarted the IIS application pool

---

## Troubleshooting

**"Could not load file or assembly 'MySql.Data'"** — Place `MySql.Data.dll` (version 8.0.30) in your TaskListAI `Bin` folder.

**Connection timeout errors** — Increase `executionTimeout` in `<httpRuntime>` (default is 36000 seconds = 10 hours).

**Large file uploads failing** — Increase `maxRequestLength` in `<httpRuntime>` (default is 8192 KB = 8 MB, maximum is 2147483647).

**Emails not sending** — Verify the email and password match in both `<appSettings>` and `<system.net>`. For Gmail, use an App Password.

**IIS 500 errors** — Check that `debug="true"` is set in `<compilation>` to see detailed errors. Verify `validateIntegratedModeConfiguration="false"` is present in `<system.webServer>`.

**Wrong column/table names in queries** — Adjust `ourdbcase`, `csvdbcase`, and `userdbcase` to match your database's case sensitivity rules.

