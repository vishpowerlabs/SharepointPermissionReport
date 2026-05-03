# SharePoint Permission Viewer Web Part

🌟 **If you find this project useful, please consider giving it a star!** 🌟 
💡 **Have suggestions or new ideas? We'd love to hear from you—please open an issue or start a discussion!**

Welcome to the SharePoint Permission Viewer Web Part! This essential administrative tool provides a comprehensive, "Single Pane of Glass" view for all site security. It helps Site Owners and Administrators manage broken inheritance, identify orphaned users, and audit site permissions efficiently.

## ✨ Key Features

- **Centralized Dashboard**: View statistics cards for Total Users, SharePoint Groups, Unique Permissions, and Orphaned Users. Easily switch between Site Permissions and Lists & Libraries.
- **Visual Permission Analysis**: Color-coded badges for permission levels (Full Control, Edit/Contribute, Read) and clear indicators for inherited vs. unique permissions.
- **Deep Clean & Group Management**: Identify and remove orphaned users (disabled/deleted in Azure AD). Expand SharePoint groups and remove members directly from the interface.
- **Deep Scan**: Validate every single item in a list or library to catch broken inheritance at the file or folder level.
- **Check Access (Audit View)**: Search for any user using the People Picker to see their explicit permissions and run targeted deep scans to ensure they have no unexpected access.
- **Theme Awareness**: Automatically adapts to your SharePoint site's active theme, including dark mode and high-contrast support.
- **Comprehensive Exports**: Export site-level, list-level, and deep scan permission reports to CSV for offline auditing.

## 🛠 Toolchain Info

This project leverages the modern SharePoint Framework (SPFx) toolchain with Heft for fast and efficient builds.

- **Node.js**: `>=22.14.0 < 23.0.0`
- **SPFx Version**: `1.22.0`
- **React**: `17.0.1`
- **TypeScript**: `~5.8.0`
- **Package Manager**: NPM
- **Build System**: Heft (`@rushstack/heft` `1.1.2`)

### Build & Deploy Instructions

1. **Install Dependencies**:
   ```bash
   npm install
   ```
2. **Build & Package for Production**:
   ```bash
   npm run build
   ```
   *This executes `heft build --clean --production` and `heft package-solution --production`.*

3. **Local Testing**:
   ```bash
   npm start
   ```

4. **Deployment**:
   Locate the generated `.sppkg` file in `sharepoint/solution` and upload it to your Site Collection App Catalog.

## ⚙️ Web Part Property Configuration

The web part is highly customizable to perfectly fit your site's needs. Use the property pane to configure the following settings:

| Property Name                                       | Type         | Description                                                                                                   |
| :-------------------------------------------------- | :----------- | :------------------------------------------------------------------------------------------------------------ |
| **Show Web Part Header**                            | Toggle       | Toggle the visibility of the "Permission Viewer" header block.                                                |
| **Show Statistics** (`showStats`)                   | Toggle       | Toggle visibility of summary statistics cards (Total Users, Groups, Broken Inheritance).                      |
| **Header Opacity** (`headerOpacity`)                | Slider       | Adjust the visual style and transparency of the header background.                                            |
| **Font Sizes**                                      | Dropdowns    | Independently select custom sizes for the Web Part Title, Content (Table headers & rows), and action Buttons. |
| **Storage Format** (`storageFormat`)                | Dropdown     | Specify the format for storage reporting display (`Auto`, `MB`, `GB`, `TB`).                                  |
| **Excluded Lists** (`excludedLists`)                | Multi-select | Exclude specific system lists (e.g., 'Site Assets', 'Microfeed') to filter out noise from the reports.        |
| **Simulate Access Denied** (`simulateAccessDenied`) | Toggle       | Force an "Access Denied" state to safely test administrative error modes and UI behaviors.                    |

## 🔒 Security & Architecture

- **Context-Aware**: Runs in the context of the currently logged-in user. It relies on SharePoint's native security trimming.
- **No Data Persistence**: All permission data is fetched via REST API on-the-fly and processed in-memory client-side.
- **SonarQube A-Rated**: Passed strict static analysis with 0 security issues and 0 bugs.

## 📖 Learn More

For a comprehensive dive into all features—including the Deep Clean (Orphaned User Scanner), Check Access Audit Tools, and Theme Awareness capabilities:

**👉 [Visit our Blog for the Full Feature Breakdown](https://www.wrvishnu.com/sharepoint-permission-viewer/?utm_source=github&utm_medium=readme&utm_campaign=sharepoint_permission_viewer)**

*(See `blog.md` and `presentation.md` in this repository for offline feature documentation and architecture details.)*
