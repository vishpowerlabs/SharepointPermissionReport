# Publishing Guide: SharePoint Permission Viewer

This guide outlines the steps required to publish the **Permission Viewer** web part to the Microsoft Commercial Marketplace (AppSource).

## 1. Pre-requisites

Before starting, ensure you have:
- A [Microsoft Partner Center](https://partner.microsoft.com/) account.
- Enrolled in the **Microsoft 365 App Publisher** program.
- A **MPN ID** (Microsoft Partner Network ID).

## 2. Prepare Your Solution Package

You must update the `config/package-solution.json` file with accurate metadata before building the final package.

### 2.1. Update Metadata
Open `config/package-solution.json` and fill in the following fields:

- **`solution.version`**: Ensure this is `1.0.0.0` (or increment if updating).
- **`developer`**:
    - `name`: Your Company Name or Publisher Name.
    - `websiteUrl`: Your company website.
    - `privacyUrl`: URL to your Privacy Policy (Required for AppSource).
    - `termsOfUseUrl`: URL to your Terms of Use (Required for AppSource).
    - `mpnId`: **CRITICAL**. Replace `Undefined-1.18.2` with your actual MPN ID.
- **`metadata`**:
    - `shortDescription`: A distinctive short description.
    - `longDescription`: A detailed description (you can reuse the "Challenge" and "Solution" sections from `blog.md`).
    - `screenshotPaths`: Add paths to valid screenshot files if you want them embedded in the package (though AppSource often asks for them separately).

### 2.2. Create Production Build
Run the following commands to create the release package:

```bash
npm run build
```
*(This runs `heft build --clean --production` and `heft package-solution --production`)*

**Output File**: `sharepoint/solution/permission-viewer.sppkg`

## 3. Create Offer in Partner Center

1.  Log in to **Partner Center**.
2.  Go to **Marketplace offers** > **Commercial Marketplace**.
3.  Click **+ New offer** > **SharePoint solution**.
4.  **Offer ID**: Create a unique ID (e.g., `permission-viewer`).
5.  **Offer Alias**: Internal name (e.g., `Permission Viewer Web Part`).

## 4. Define Offer Properties

- **Categories**: Select relevant categories (e.g., *Collaboration*, *Productivity*, *IT & Administration*).
- **Industries**: Select if applicable (or "None").
- **App version**: Must match `1.0.0.0` from your `package-solution.json`.

## 5. Offer Listing Details

This is what customers see in the store. You can use content from `blog.md` and `linkedin_post.md` here.

- **Name**: `SharePoint Permission Viewer`
- **Short Description**: `Centralized dashboard to view, audit, and manage SharePoint permissions.`
- **Description**: Use the HTML editor to paste your full value proposition.
- **Search Keywords**: `SharePoint`, `Permissions`, `Security`, `Audit`, `Governance`.
- **Products your app works with**: Select **SharePoint**.
- **Screenshots**: Upload at least one screenshot (1280x720 or 1366x768). You can use `assets/permission_report_output.png`.
- **Store Logos**:
    - **Large**: 216x216 px (PNG)
    - **Small**: 48x48 px (PNG)
    - *Note: You may need to resize `assets/blog_feature_image.png` to these dimensions.*

## 6. Technical Configuration

1.  **File Upload**: Drag and drop your `permission-viewer.sppkg` file here.
2.  **Validation**: The system will parse the package. It checks if the `mpnId` in `package-solution.json` matches your Partner Center account.

## 7. Availability

- **Markets**: Select the countries where you want to sell/distribute the app.
- **Pricing**: Choose **Free** or **Transact** (if you plan to charge).

## 8. Review and Publish

1.  Click **Review and publish**.
2.  Review the certification notes.
3.  Click **Publish**.

### Certification Process
Microsoft will validate your app:
1.  **Automated Validation**: Checks for viruses, malware, and package errors.
2.  **Certification**: A manual review by Microsoft testers. They may ask for a test account or demo environment if your app requires complex setup.
    - *Since this web part runs on the user's data, you usually don't need to provide creds, but you must explain how to test it in the "Notes for certification" (e.g., "Add the web part to any SharePoint site").*

**Timeline**: Certification typically takes 1-3 business days.
