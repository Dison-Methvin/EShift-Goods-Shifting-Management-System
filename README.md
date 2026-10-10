# E_Shift_V02

## Overview

E_Shift_V02 is a Windows Forms-based transport job management system that centralizes registration, job management, real-time updates, and automated reporting for administrators and customers.

## Scenario

The system addresses key challenges such as inefficient job tracking, poor communication, and manual data handling by offering a centralized digital solution. It comprises dedicated modules for both administrators and customers, enabling seamless registration, transport job management, real-time job updates, and automated report generation.

Key technologies used in the development include C# with Windows Forms, SQL Server, and Leaflet.js for map rendering. The system incorporates Object-Oriented Programming (OOP) principles, robust validation mechanisms, and secure data handling through SHA256 encryption. Features such as PDF report exports, email notifications, and location-based visualizations enhance user experience and operational visibility.

## Key Features
- Admin and Customer modules (separate UI flows)
- User registration and authentication
- Transport job creation, assignment, and status tracking
- Real-time job updates and location visualization (Leaflet.js)
- PDF report export (iTextSharp)
- Email notifications
- Secure data handling using SHA256 hashing for sensitive values
- Exportable reports and automated reporting

## Tech Stack
- Language: C# (.NET / Windows Forms)
- UI framework: Windows Forms
- Database: Microsoft SQL Server
- Mapping: Leaflet.js (for map rendering in web-enabled components)
- PDF generation: iTextSharp
- JSON: Newtonsoft.Json
- Other: WebView2 (where applicable), various System.* packages

## Architecture & File Structure

Top-level layout (workspace root and main project folder):

- E_Shift_V02.sln
- E_Shift_V02/
  - App.config
  - Program.cs
  - Admin_Dashboard.cs
  - Admin_Login.cs
  - Customer_Dashboard.cs
  - Customer_Login.cs
  - Customer_SignUp.cs
  - Role_Selection.cs
  - Managers/ (Agent managers and business logic)
    - AssistantManager.cs
    - ContainerManager.cs
    - DriverManager.cs
    - LorryManager.cs
    - MailingManager.cs
    - ProductManager.cs
  - Models/
    - User.cs
  - Services/
    - AuthService.cs
    - UserService.cs
  - Properties/
    - AssemblyInfo.cs
    - Resources.resx
    - Settings.settings
  - bin/, obj/
- packages/ (local NuGet packages used by the project)

Note: The exact file listing may include designer and resource files (`*.Designer.cs`, `*.resx`) for each WinForms form.

## Requirements
- Windows with Visual Studio (2019/2022 recommended) with .NET desktop development workload
- Microsoft SQL Server (or SQL Server Express)
- Restore NuGet packages (solution uses packages in `packages/` and may reference online packages)

## Setup & Run (Developer)

1. Clone the repository or extract the project folder.
2. Open the solution `E_Shift_V02.sln` in Visual Studio.
3. Restore NuGet packages (Visual Studio will normally restore on open). If needed, run `nuget restore` or use the NuGet UI.
4. Configure the database:
   - Create a SQL Server database for the application (example name: `E_Shift_V02_DB`).
   - Update the connection string in `App.config` to point to your SQL Server instance and database.
   - If you have SQL initialization scripts, run them to create tables and seed data. (If no scripts exist, create the schema according to the models in `Models/` or ask the maintainers for the DDL.)
5. Build the solution (Build → Build Solution).
6. Run the project (F5) — the Windows Forms application will launch.

Example `App.config` connection string entry (update with your server/name/password):

```xml
<connectionStrings>
  <add name="EShiftDb" connectionString="Server=localhost;Database=E_Shift_V02_DB;User Id=sa;Password=your_password;" providerName="System.Data.SqlClient" />
</connectionStrings>
```

## Database notes
- The app expects a SQL Server database. Check `App.config` for the connection key used by `Services/` classes.
- If you don't have DB scripts, inspect `Services/` and `Models/` to derive schema and create the tables manually.

## Security
- The application uses SHA256 hashing for sensitive values — ensure any secrets (connection strings, SMTP credentials) are stored securely when deploying (use environment variables or a secure secret store in production).

## Packaging & Reports
- PDF export uses iTextSharp; generated reports are usually saved or streamed to the user via the UI.

## Testing
- Manual testing of user flows (Admin & Customer) is expected since this is a WinForms app.
- Verify: registration, login, job creation, job updates, PDF export, and email notifications.
