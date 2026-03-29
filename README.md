# SharePoint Permission Viewer Web Part

Welcome to the SharePoint Permission Viewer Web Part! This essential administrative tool provides a comprehensive, "Single Pane of Glass" view for all site security. It helps Site Owners and Administrators manage broken inheritance, identify orphaned users, and audit site permissions efficiently.

## 🛠 Toolchain Info

This project leverages the modern SharePoint Framework (SPFx) toolchain with Heft for fast and efficient builds.

- **Node.js**: `v22.14.0` or higher (requires `< 23.0.0`)
- **SPFx Version**: `1.22.0`
- **React**: `17.0.1`
- **TypeScript**: `~5.8.0`
- **Package Manager**: NPM
- **Build System**: Heft (`@rushstack/heft`)

### Build & Deploy Instructions

1.  **Install Dependencies**:
    ```bash
    npm install
    ```
2.  **Build & Package for Production**:
    ```bash
    npm run build
    ```
    *This executes `heft build --clean --production` and `heft package-solution --production`.*

3.  **Local Testing**:
    ```bash
    npm start
    ```

## ⚙️ Web Part Property Configuration

The web part is highly customizable to perfectly fit your site's needs. Use the property pane to configure the following settings:

| Property Name | Type | Description |
| :--- | :--- | :--- |
| **Show Web Part Header** | Toggle | Toggle the visibility of the "Permission Viewer" header block. |
| **Show Statistics** (`showStats`) | Toggle | Toggle visibility of summary statistics cards (Total Users, Groups, Broken Inheritance). |
| **Header Opacity** (`headerOpacity`) | Slider | Adjust the visual style and transparency of the header background. |
| **Font Sizes** | Dropdowns | Independently select custom sizes for the Web Part Title, Content (Table headers & rows), and action Buttons. |
| **Storage Format** (`storageFormat`) | Dropdown | Specify the format for storage reporting display (`Auto`, `MB`, `GB`, `TB`). |
| **Excluded Lists** (`excludedLists`) | Multi-select | Exclude specific system lists (e.g., 'Site Assets', 'Microfeed') to filter out noise from the reports. |
| **Simulate Access Denied** (`simulateAccessDenied`) | Toggle | Force an "Access Denied" state to safely test administrative error modes and UI behaviors. |

## 📖 Learn More

For a comprehensive dive into all features—including the Deep Clean (Orphaned User Scanner), Check Access Audit Tools, and Theme Awareness capabilities:

**👉 [Visit our Blog for the Full Feature Breakdown](https://vishpowerlabs.com/blog/sharepoint-permission-viewer)**

*(See `blog.md` and `presentation.md` in this repository for offline feature documentation and architecture details.)*
