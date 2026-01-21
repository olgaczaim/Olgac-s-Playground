# SharePoint FBA Pack Documentation

Detailed documentation for installing, configuring, and using the **SharePoint FBA Pack**.

> The SharePoint FBA Pack is a set of tools and web parts that provide user management and self-service features for SharePoint sites using Forms Based Authentication (FBA). It supports SharePoint 2010, 2013, 2016, and 2019.

---

## 📦 Installing and Configuring the SharePoint FBA Pack

### 1. Configure SharePoint to Use Forms Based Authentication

Before installing the FBA Pack, SharePoint must be configured to use Forms Based Authentication.

Depending on your SharePoint version, follow the appropriate online guide:

- **SharePoint 2010:** [Setup FBA using IIS and SQL Membership Provider](https://donalconlon.wordpress.com/2010/02/23/configuring-forms-base-authentication-for-sharepoint-2010-using-iis7/)
- **SharePoint 2013:** [Configure FBA and create the membership DB](http://blogs.visigo.com/chriscoulson/configuring-forms-based-authentication-in-sharepoint-2013-part-1-creating-the-membership-database/)
- **SharePoint 2016/2019:** [Similar FBA configuration for newer versions](http://blogs.visigo.com/chriscoulson/configuring-forms-based-authentication-in-sharepoint-2016-part-1-creating-the-membership-database/)

([Links referenced in the official guide](https://www.visigo.com/products/sharepoint-fba-pack/documentation/))

#### Configuring Forms Based Authentication (FBA) in SharePoint 2016 / 2019 — Part 1: Creating the Membership Database

##### 🛠 Step 1 — Run the ASP.NET Membership Database Wizard

1. Open **Windows Explorer** and navigate to this folder on the SharePoint server:

   ```
   C:\Windows\Microsoft.NET\Framework64\v4.0.30319\
   ```

2. Find and double-click:

   ```
   aspnet_regsql.exe
   ```

   This launches the **ASP.NET SQL Server Setup Wizard**.

3. Click **Next** on the welcome screen.

4. Choose the option:

   ```
   Configure SQL Server for application services
   ```

5. Enter the name of your server and your authentication information.  In this case SQL Server is installed on the same server as SharePoint and I am logged in as an administrator and have full access to SQL Server, so I choose Windows Authentication.For the database name, I just leave it as <default>, which creates a database called “aspnetdb”.
6. A Confirm Your Settings screen will appear. Click Next.
7. A “database has been created or modified” screen will appear. Click finish and the wizard will close.
8. Now that we know what account is being used to run SharePoint, we can assign it the appropriate permissions to the membership database we created.  Open up SQL Server Management Studio and log in as an administrator.
9. Under Security/Logins find the user that SharePoint runs as.  Assuming this is the same database server that SharePoint was installed on, the user should already exist.Right click on the user and click ‘Properties’.
10. Go to the “User Mapping” Page. Check the “Map” checkbox for the aspnetdb database. With the aspnetdb database selected, check the “db_owner” role membership and click OK. This user should now have full permissions to read and write to the aspnetdb membership database.

##### 🛠 Step 2 — Editing the Web.Config Files

**Option A — Machine.config (preferred):**

- When settings are placed in the `machine.config`, all web applications inherit the FBA provider settings.
- This prevents repeated edits for each new SharePoint web app.  
- Applies to: Central Admin, web apps, STS, etc.  
- **Important:** If you have multiple servers (farm), repeat on all.  
- Always **backup any .config file before editing**.

**Option B — Individual Web.config Files:**

- Only use if you cannot edit the machine.config.
- You will then need to apply these edits to:
  - Central Administration
  - Security Token Service
  - Each SharePoint web application you want to support FBA.


**1. Open the machine.config file**

Usually located at:

```
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Config\machine.config
```  
Open in a text editor (with administrator rights).

**2. Add Connection String**

In the `<connectionStrings>` section, add:

```xml
<add name="FBADB"
     connectionString="Server=<SQL_SERVER_NAME>;
                       Database=aspnetdb;
                       Integrated Security=true" />
```

Replace `<SQL_SERVER_NAME>` with your actual SQL instance name.

**3. Add Membership Provider**

Inside `<system.web><membership><providers>` add:

```xml
<add name="FBAMembershipProvider"
     type="System.Web.Security.SqlMembershipProvider, System.Web, Version=4.0.0.0,
          Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
     connectionStringName="FBADB"
     enablePasswordRetrieval="false"
     enablePasswordReset="true"
     requiresQuestionAndAnswer="false"
     applicationName="/"
     requiresUniqueEmail="true"
     passwordFormat="Hashed"
     maxInvalidPasswordAttempts="5"
     minRequiredPasswordLength="7"
     minRequiredNonalphanumericCharacters="1"
     passwordAttemptWindow="10"
     passwordStrengthRegularExpression="" />
```

> The critical part is using the same options everywhere the provider is defined.

**4. Add Role Provider**

Inside `<system.web><roleManager><providers>` add:

```xml
<add name="FBARoleProvider"
     type="System.Web.Security.SqlRoleProvider, System.Web, Version=4.0.0.0,
          Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
     connectionStringName="FBADB"
     applicationName="/" />
```

Save and close the file.

---

###### 🔁 Special Case — SecurityTokenService Web.config

Even after editing `machine.config`, you must also edit the **SecurityTokenServiceApplication web.config** because its settings *override* the machine.config entries.

**1. Find the SecurityTokenService Application**

- Open **IIS Manager**
- Under **Sites → SharePoint Web Services**
- Right-click **SecurityTokenServiceApplication** → *Explore*
- Open the `web.config` in that folder.

**2. Add Membership and Role Providers**

Before the closing `</configuration>` tag, add:

```xml
<system.web>
  <membership>
    <providers>
      <add name="FBAMembershipProvider"
           type="System.Web.Security.SqlMembershipProvider, System.Web,
                Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
           connectionStringName="FBADB"
           enablePasswordRetrieval="false"
           enablePasswordReset="true"
           requiresQuestionAndAnswer="false"
           applicationName="/"
           requiresUniqueEmail="true"
           passwordFormat="Hashed"
           maxInvalidPasswordAttempts="5"
           minRequiredPasswordLength="7"
           minRequiredNonalphanumericCharacters="1"
           passwordAttemptWindow="10"
           passwordStrengthRegularExpression="" />
    </providers>
  </membership>

  <roleManager enabled="true">
    <providers>
      <add name="FBARoleProvider"
           type="System.Web.Security.SqlRoleProvider, System.Web,
                Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
           connectionStringName="FBADB"
           applicationName="/" />
    </providers>
  </roleManager>
</system.web>
```

Save and close the file.

---

###### 🛡️ Providers Explained

| Setting | Purpose |
|---------|---------|
| `connectionStringName` | Which database to use for FBA. |
| `enablePasswordRetrieval` | Whether plain passwords can be retrieved. (usually false) |
| `enablePasswordReset` | Allow users to reset passwords. |
| `requiresQuestionAndAnswer` | Whether QA is required for reset. |
| `passwordFormat` | Hashed is most secure. |
| `applicationName` | Scope of users in the database. |
| `roleManager` settings | Connects role provider to the same database. | 

---

###### 🧠 Important Notes

- All changes must be applied on **every SharePoint server** in the farm. 
- Always **back up** config files before editing. 
- Make sure the membership and role provider names and options are consistent wherever they’re defined.

---

##### 🛠 Step 3 — Configuring SharePoint

1. Open **SharePoint Central Administration**.
2. Go to **Application Management** → **Manage Web Applications**.
3. Select the web application you want to enable FBA on.
4. Click **Authentication Providers** on the ribbon.
5. Choose the zone you wish to configure (e.g., **Default**).
6. On the Authentication Provider screen:
   - **Check** the box for **Enable Forms Based Authentication (FBA)**.
   - In the **ASP.NET Membership Provider Name** field enter your membership provider (e.g., `FBAMembershipProvider`).
   - In the **ASP.NET Role Manager Name** field enter your role provider (e.g., `FBARoleProvider`).
   - Optionally leave **Enable Windows Authentication** checked to allow both Windows and FBA logins.
7. Click **Save**. 

After saving, users attempting to access the site will see both authentication options (if Windows Auth is enabled) or just FBA if Windows Auth is disabled.

---

## 💡 Notes

- If **both Windows and Forms Authentication** are enabled, SharePoint will prompt users to choose which method to use when signing in.
- At this stage, FBA may still not allow logins until users have been **added to the membership database**, which is covered in **Part 4** of the series.

---

## ⚠️ Special Case — Office Clients and Modern Authentication

For Office applications (like Word or Excel) to authenticate properly with FBA, there’s a known issue when **Modern Authentication** is enabled by default in SharePoint 2016.  
To support legacy Claims authentication in Office clients, the following PowerShell must be run:

```powershell
$sts = Get-SPSecurityTokenServiceConfig
$sts.SuppressModernAuthForOfficeClients = $True
$sts.Update()
iisreset
```

This disables Modern Authentication for Office clients so they can authenticate via traditional claims (required for FBA to work with Office apps).

---


---

##### 🛠 Step 4 — Adding Users to the Membership Database

1. Download the **SharePoint FBA Pack** for your SharePoint version from:

   ```
   https://www.visigo.com/products/sharepoint-fba-pack
   ```

2. Unzip the package to a folder, e.g.:

   ```
   C:\deploy
   ```

3. Open **PowerShell** (run as Administrator) or the **SharePoint Management Shell**.
4. Change directory to where you unzipped the FBA Pack:

   ```powershell
   cd C:\deploy
   ```

5. Run the deploy script:

   ```powershell
   .\deploy http://<YourSiteCollectionURL>
   ```

   - Replace `<YourSiteCollectionURL>` with the URL of the site collection where you want FBA management enabled.
   - If you run `.\deploy` *without* a URL, you must manually activate the **Forms Based Authentication Management** feature in each site collection you need FBA in.

6. If PowerShell blocks script execution, run:

   ```powershell
   Set-ExecutionPolicy Unrestricted
   ```

   then re-run the deploy script.

7. Navigate to your site collection in SharePoint.
8. Log in as a **Site Collection Administrator**.
9. Go to **Site Settings**.
10. You should now see additional FBA options, including:

   - **FBA User Management**
   - **FBA Role Management** (if roles are enabled)

11. Click **FBA User Management**.
12. Click **New User**.
13. Fill in the form to create a new FBA user:

   - Username
   - Password
   - Email
   - (Optional) Role assignment
   - (Optional) Membership review settings — if your FBA site configuration requires admin review, new users may need approval.

14. Assign the FBA user to a SharePoint group so they have permissions to actually use the site.
15. Log out of SharePoint.
16. Go to your site’s login page.
17. Choose **Forms Based Authentication**.
18. Enter the FBA username and password you just created.

You should now be able to log in using the FBA account.

## 💡 Notes & Tips

- **Permissions:** Make sure the application pool identity you use for your SharePoint web application has **db_owner** rights on the membership database. Otherwise management pages won’t work and logins may fail.
- **Troubleshooting:** If you get **“A Membership Provider has not been configured correctly”**, double-check your *web.config* and *SecurityTokenService* config entries for the membership providers
- **Self-Service:** In addition to user creation in the FBA Pack, there are web parts for **user registration**, **password change**, and **password recovery** if needed.

---
