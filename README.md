# Azure Automation Runbooks

A comprehensive collection of PowerShell runbooks for Azure, Intune, and Entra ID automation. This repository includes scripts for generating detailed reports on costs, MFA status, license availability, PIM activations, and more. All scripts are designed to run in **Azure Automation** using **PowerShell 7.2**.

## 🚀 Features

-   **Automated Reporting**: Generate HTML reports for Azure costs, Entra ID status, and Intune compliance.
-   **Email Integration**: Automatically send reports via email using Microsoft Graph API.
-   **Visual Alerts**: Reports include color-coded alerts for critical items (e.g., expiring secrets, inactive users).
-   **Secure**: Uses Azure Automation Variables and Credentials for secure secret management.

## 📂 Repository Structure

| Runbook Name | Description |
| :--- | :--- |
| **Runbook_Azure-Groups_DynamicGroupByPropertiesTenantFree.ps1** | Synchronizes an Azure AD group by adding/removing users based on `EmployeeId`. |
| **Runbook_Azure-Licenses_GroupBasedLicenseTenantFree.ps1** | Assigns the *Fabric Free* license to users who already hold *Power BI Pro*. |
| **Runbook_Azure-Report_AppRegistrationsExpirationCertificates.ps1** | Generates a report of App Registration certificates nearing expiration. |
| **Runbook_Azure-Report_AppRegistrationsExpirationSecrets.ps1** | Generates a report of App Registration secrets nearing expiration. |
| **Runbook_Azure-Report_MultiTenantCosts.ps1** | Generates a monthly cost report (incl. VAT) for subscriptions across **two** tenants. |
| **Runbook_Azure-Report_TenantCosts.ps1** | Generates a monthly cost report (incl. VAT) for subscriptions in a **single** tenant. |
| **Runbook_EntraID-Report_AvailableLicenses.ps1** | Detailed report of available Microsoft 365 licenses, highlighting shortages. |
| **Runbook_EntraID-Report_DisableUsersInactivty.ps1** | Checks for inactive users and can optionally block them or run in simulation mode. |
| **Runbook_EntraID-Report_MFA_Disabled.ps1** | Identifies users who do **not** have MFA enabled. |
| **Runbook_EntraID-Report_MFA_Enabled_Disabled.ps1** | Overview of MFA status (Enabled/Disabled) for all users. |
| **Runbook_EntraID-Report_PIMActivationJustification.ps1** | Audits PIM activations and highlights justifications that are too short or missing. |
| **Runbook_Intune-Report_NonCompliantDevices.ps1** | Reports non-compliant Intune devices and optionally notifies the user. |
| **Runbook_Sharepoint-Report_LargerThan500MB.ps1** | Scans SharePoint sites for files larger than 500MB. |

## 🛠️ Prerequisites

-   **Azure Automation Account** with System Assigned Identity or App Registration.
-   **PowerShell 7.2 Runtime** enabled in Azure Automation.
-   **Microsoft Graph API Permissions**: Each script header documents the specific permissions required (e.g., `User.Read.All`, `Mail.Send`, `AuditLog.Read.All`).
-   **Modules**: Ensure `Microsoft.Graph` modules are imported in the Automation Account.

## ⚙️ Configuration

1.  **Variables**: Create the necessary Variables in Azure Automation (e.g., `GraphClientId`, `GraphTenantId`, `GraphSecret`).
2.  **Credentials**: Create a Credential asset (e.g., `Correo_No-Reply`) for sending emails.
3.  **Schedules**: Link the runbooks to Schedules for automated execution.

## 📝 License

This project is open-source and available for use and modification.

---