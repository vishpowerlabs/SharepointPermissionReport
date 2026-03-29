# SharePoint Permission Viewer - Feature Guide

## Overview
The **SharePoint Permission Viewer** is a powerful administrative tool designed to give Site Owners and Administrators complete visibility and control over their SharePoint site's security. It goes beyond standard SharePoint settings by providing deep insights into broken inheritance, orphaned users, and comprehensive access reports.

---

## 🛡️ Key Features

### 1. 📊 Security & Governance Dashboard
**What it does:**
- Provides an immediate high-level overview of the site's security posture.
- Metrics include:
    - **Total Users**: Count of unique users with access.
    - **SharePoint Groups**: Number of SP groups defined.
    - **Broken Inheritance**: **Critical** metric showing how many lists/items have unique permissions.
    - **Orphaned Users**: Users who still have permission records but are disabled or deleted in Active Directory.

**Why it's helpful:**
- Instantly spot security anomalies (e.g., a sudden spike in broken permissions).
- Track the "cleanliness" of your site's user base.

### 2. 🔍 Check Access (User Access Report)
**What it does:**
- Allows you to search for **any specific user** and see *exactly* what they can access across the entire site.
- **Scope**: Checks Site level, List/Library level, and even individual Items/Folders with unique permissions.
- **Export**: Generates a CSV report of that user's access.

**Why it's helpful:**
- **Onboarding/Offboarding**: Verify a new joiner has the right access, or confirm a leaver has been revoked.
- **Troubleshooting**: Quickly answer "Why can't User X see this file?" or "Why can User Y see this confidential list?"
- **Auditing**: Provide proof of access for compliance reviews.

### 3. 🧹 Deep Clean (Orphaned User Scanner)
**What it does:**
- Performs a **comprehensive scan** of the entire site (Lists, Libraries, Folders, Items).
- Identifies **"Orphaned Users"**: Accounts that are Disabled or Deleted in Azure AD but still hold permissions on SharePoint items.
- **Actionable**: Allows you to **remove** these invalid permissions directly from the interface.

**Why it's helpful:**
- **Security Logic**: "Deleted" users can't login, but their permission entries clutter the ACLs (Access Control Lists) and can cause technical issues or confusion.
- **Performance**: Excessive "dead" permissions can degrade SharePoint performance.
- **Compliance**: Ensures your permission reports only reflect active, valid users.

### 4. 👥 Site Groups & Members
**What it does:**
- Displays all SharePoint Groups (Owners, Members, Visitors, etc.) and their membership.
- **Orphan Indicators**: visually flags users within groups who are disabled/deleted.
- **Member Management**: View full details of group members.

**Why it's helpful:**
- Quickly audit who is in the "Owners" group (privileged access).
- Clean up groups without navigating to the backend settings pages.

### 5. 📂 List & Library Permission Overview
**What it does:**
- Lists all Libraries and Lists on the site.
- **Visual Badge**: Instantly shows if a list has **"Unique Permissions"** (Broken Inheritance).
- **Drill-down**: Click to see exactly *who* has access to that specific list.

**Why it's helpful:**
- **Broken Inheritance** is the #1 cause of permission management nightmares. This view highlights exactly where inheritance is broken so you can manage it.
- Prevents "hidden" access where a sub-folder might be accessible to users who don't have access to the parent library.

---

## 🚀 Best Practices & Use Cases

### Scenario A: Employee Leaving the Company
1.  **Run "Deep Clean"**: Scan the site to find any direct permissions assigned to their account on specific files or folders.
2.  **Remove Permissions**: Use the clean-up tool to strip their account from those items.
3.  **Check Groups**: Go to Site Groups and remove them from standard groups (Members/Visitors).
4.  **Verify**: Run "Check Access" for their name to ensure it returns "No Access Found".

### Scenario B: Preparing for a Security Audit
1.  **Review Dashboard**: Ensure "Orphaned Users" is zero.
2.  **Review "Unique Permissions"**: Check lists with broken inheritance. Ask: *Does this list really need unique permissions, or should it inherit from the site?*
3.  **Export Reports**: Use the "Check Access" export for key sensitive users (e.g., External Auditors, Guests).

---

## 🛠️ Technical Architecture

### API Usage
The web part is built for **performance** and **minimal dependency**:

*   **SharePoint REST API**: Used for **90% of operations**, including:
    - Lists, Permissions, and User retrieval.
    - User Profile Service (UPS) for user details.
    - Search API for "Check Access".

This ensures fast, native performance for standard tasks while leveraging the power of SharePoint native APIs without requiring high-privilege Graph consents.

---

## ⚠️ Performance & Limitations

### Large Site Considerations
While efficient, this tool runs **client-side** in your browser. It is **not recommended** for:
*   **Extremely Large Sites**: Sites with hundreds of thousands of items with unique permissions.
*   **Deeply Nested Structures**: Libraries with excessive folder depth (>10 levels) and broken inheritance at every level.

**Why?**
Scanning every single item for unique permissions in a massive list can lead to:
*   **Browser Throttling**: The browser may become unresponsive.
*   **API Throttling**: SharePoint may temporarily block requests (Error 429) to protect the server.

**Recommendation:**
For massive enterprise migrations or audits of sites with >100,000 unique permissions, consider using a server-side script (PowerShell) or a dedicated third-party governance tool.
