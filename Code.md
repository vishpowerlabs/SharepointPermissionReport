<!DOCTYPE html>
<html>
<head>
<style>
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    margin: 0;
    padding: 0;
    background: #faf9f8;
}

.webpart-container {
    max-width: 1200px;
    margin: 20px auto;
    background: #ffffff;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    overflow: hidden;
}

.header {
    background: linear-gradient(90deg, #0078d4 0%, #6b69d6 100%);
    padding: 24px 32px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.header h1 {
    color: #ffffff;
    font-size: 24px;
    font-weight: 600;
    margin: 0;
}

.refresh-btn {
    background: rgba(255,255,255,0.2);
    border: 1px solid rgba(255,255,255,0.3);
    color: #ffffff;
    padding: 8px 16px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
}

.refresh-btn:hover {
    background: rgba(255,255,255,0.3);
}

.stats-container {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 20px;
    padding: 24px 32px;
    background: #faf9f8;
}

.stat-card {
    background: #ffffff;
    border-radius: 8px;
    padding: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    transition: all 0.2s ease;
}

.stat-card:hover {
    box-shadow: 0 4px 16px rgba(0,0,0,0.15);
    transform: translateY(-2px);
}

.stat-icon {
    font-size: 24px;
    margin-bottom: 8px;
}

.stat-number {
    font-size: 32px;
    font-weight: 700;
    color: #323130;
    margin: 8px 0;
}

.stat-label {
    font-size: 14px;
    color: #605e5c;
}

.tabs-container {
    border-bottom: 1px solid #e1dfdd;
    padding: 0 32px;
}

.tabs {
    display: flex;
    gap: 32px;
}

.tab {
    padding: 16px 0;
    font-size: 16px;
    color: #605e5c;
    border-bottom: 3px solid transparent;
    cursor: pointer;
    transition: all 0.2s ease;
}

.tab:hover {
    color: #323130;
    background: #f3f2f1;
}

.tab.active {
    color: #0078d4;
    border-bottom-color: #0078d4;
    font-weight: 600;
}

.toolbar {
    padding: 16px 32px;
    background: #ffffff;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid #e1dfdd;
}

.search-box {
    flex: 1;
    max-width: 400px;
    padding: 8px 16px;
    border: 1px solid #e1dfdd;
    border-radius: 4px;
    font-size: 14px;
}

.export-btn {
    background: #0078d4;
    color: #ffffff;
    border: none;
    padding: 8px 20px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 14px;
    font-weight: 600;
}

.export-btn:hover {
    background: #106ebe;
}

.content {
    padding: 24px 32px;
}

.permission-table {
    width: 100%;
    border-collapse: collapse;
    background: #ffffff;
    border: 1px solid #e1dfdd;
    border-radius: 4px;
    overflow: hidden;
}

.permission-table thead {
    background: #faf9f8;
}

.permission-table th {
    text-align: left;
    padding: 12px 16px;
    font-size: 14px;
    font-weight: 600;
    color: #323130;
    border-bottom: 1px solid #e1dfdd;
}

.permission-table td {
    padding: 12px 16px;
    border-bottom: 1px solid #f3f2f1;
}

.permission-table tr:hover {
    background: #f3f2f1;
}

.user-cell {
    display: flex;
    align-items: center;
    gap: 12px;
}

.avatar {
    width: 32px;
    height: 32px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #ffffff;
    font-weight: 600;
    font-size: 14px;
}

.avatar-blue { background: #0078d4; }
.avatar-purple { background: #8764b8; }
.avatar-green { background: #107c10; }
.avatar-orange { background: #ca5010; }

.user-info {
    display: flex;
    flex-direction: column;
}

.user-name {
    font-size: 14px;
    color: #323130;
    font-weight: 600;
}

.user-email {
    font-size: 12px;
    color: #605e5c;
}

.badge {
    display: inline-block;
    padding: 4px 12px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid;
}

.badge-full-control {
    background: #fde7e9;
    color: #d13438;
    border-color: #d13438;
}

.badge-edit {
    background: #fef6f1;
    color: #ca5010;
    border-color: #ca5010;
}

.badge-contribute {
    background: #e6f2ff;
    color: #0078d4;
    border-color: #0078d4;
}

.badge-read {
    background: #e6f7e6;
    color: #107c10;
    border-color: #107c10;
}

.list-card {
    background: #ffffff;
    border: 1px solid #e1dfdd;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    transition: all 0.2s ease;
}

.list-card:hover {
    border-color: #0078d4;
}

.list-card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 12px;
}

.list-title {
    display: flex;
    align-items: center;
    gap: 12px;
}

.list-title h3 {
    font-size: 18px;
    font-weight: 600;
    color: #323130;
    margin: 0;
}

.list-icon {
    font-size: 20px;
}

.list-url {
    font-size: 12px;
    color: #605e5c;
    font-style: italic;
    margin-bottom: 16px;
}

.permission-badge {
    padding: 6px 16px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid;
}

.badge-unique {
    background: #fef6f1;
    color: #ca5010;
    border-color: #ca5010;
}

.badge-inherited {
    background: #f3f2f1;
    color: #605e5c;
    border-color: #8a8886;
}

.inherited-info {
    background: #e6f2ff;
    border-left: 4px solid #0078d4;
    border-radius: 4px;
    padding: 16px;
    display: flex;
    align-items: center;
    gap: 12px;
}

.inherited-info-icon {
    color: #0078d4;
    font-size: 20px;
}

.inherited-info-text {
    font-size: 14px;
    color: #323130;
}

.loading-state {
    text-align: center;
    padding: 60px 32px;
}

.spinner {
    width: 48px;
    height: 48px;
    border: 4px solid #f3f2f1;
    border-top-color: #0078d4;
    border-radius: 50%;
    animation: spin 1s linear infinite;
    margin: 0 auto 20px;
}

@keyframes spin {
    to { transform: rotate(360deg); }
}

.loading-text {
    font-size: 16px;
    color: #605e5c;
    margin: 12px 0;
}

.progress-bar {
    width: 300px;
    height: 4px;
    background: #f3f2f1;
    border-radius: 2px;
    margin: 20px auto;
    overflow: hidden;
}

.progress-fill {
    height: 100%;
    background: #0078d4;
    width: 60%;
    animation: progress 2s ease-in-out infinite;
}

@keyframes progress {
    0% { width: 0%; }
    50% { width: 60%; }
    100% { width: 100%; }
}
</style>
</head>
<body>

<!-- SITE PERMISSIONS VIEW -->
<div class="webpart-container">
    <div class="header">
        <h1>📊 SharePoint Permission Viewer</h1>
        <button class="refresh-btn">🔄 Refresh</button>
    </div>
    
    <div class="stats-container">
        <div class="stat-card">
            <div class="stat-icon" style="color: #0078d4;">👥</div>
            <div class="stat-number">15</div>
            <div class="stat-label">Total Users</div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="color: #8764b8;">👨‍👩‍👧‍👦</div>
            <div class="stat-number">8</div>
            <div class="stat-label">Groups</div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="color: #ca5010;">🔒</div>
            <div class="stat-number">3</div>
            <div class="stat-label">Unique Permissions</div>
        </div>
    </div>
    
    <div class="tabs-container">
        <div class="tabs">
            <div class="tab active">🏠 Site Permissions</div>
            <div class="tab">📚 Lists & Libraries</div>
        </div>
    </div>
    
    <div class="toolbar">
        <input type="text" class="search-box" placeholder="🔍 Search permissions...">
        <button class="export-btn">📥 Export</button>
    </div>
    
    <div class="content">
        <table class="permission-table">
            <thead>
                <tr>
                    <th>User/Group</th>
                    <th>Type</th>
                    <th>Permission Level</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>
                        <div class="user-cell">
                            <div class="avatar avatar-blue">JD</div>
                            <div class="user-info">
                                <span class="user-name">John Doe</span>
                                <span class="user-email">john.doe@company.com</span>
                            </div>
                        </div>
                    </td>
                    <td>User</td>
                    <td><span class="badge badge-full-control">Full Control</span></td>
                </tr>
                <tr>
                    <td>
                        <div class="user-cell">
                            <div class="avatar avatar-purple">SO</div>
                            <div class="user-info">
                                <span class="user-name">Site Owners</span>
                                <span class="user-email">15 members</span>
                            </div>
                        </div>
                    </td>
                    <td>SharePoint Group</td>
                    <td><span class="badge badge-full-control">Full Control</span></td>
                </tr>
                <tr>
                    <td>
                        <div class="user-cell">
                            <div class="avatar avatar-blue">SM</div>
                            <div class="user-info">
                                <span class="user-name">Site Members</span>
                                <span class="user-email">45 members</span>
                            </div>
                        </div>
                    </td>
                    <td>SharePoint Group</td>
                    <td><span class="badge badge-contribute">Contribute</span></td>
                </tr>
                <tr>
                    <td>
                        <div class="user-cell">
                            <div class="avatar avatar-green">SV</div>
                            <div class="user-info">
                                <span class="user-name">Site Visitors</span>
                                <span class="user-email">120 members</span>
                            </div>
                        </div>
                    </td>
                    <td>SharePoint Group</td>
                    <td><span class="badge badge-read">Read</span></td>
                </tr>
                <tr>
                    <td>
                        <div class="user-cell">
                            <div class="avatar avatar-orange">MT</div>
                            <div class="user-info">
                                <span class="user-name">Marketing Team</span>
                                <span class="user-email">marketing@company.com</span>
                            </div>
                        </div>
                    </td>
                    <td>Security Group</td>
                    <td><span class="badge badge-edit">Edit</span></td>
                </tr>
            </tbody>
        </table>
    </div>
</div>

<br><br>

<!-- LISTS & LIBRARIES VIEW -->
<div class="webpart-container">
    <div class="header">
        <h1>📊 SharePoint Permission Viewer</h1>
        <button class="refresh-btn">🔄 Refresh</button>
    </div>
    
    <div class="stats-container">
        <div class="stat-card">
            <div class="stat-icon" style="color: #0078d4;">📚</div>
            <div class="stat-number">12</div>
            <div class="stat-label">Total Lists</div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="color: #8764b8;">📁</div>
            <div class="stat-number">5</div>
            <div class="stat-label">Doc Libraries</div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="color: #ca5010;">🔒</div>
            <div class="stat-number">3</div>
            <div class="stat-label">Unique Permissions</div>
        </div>
    </div>
    
    <div class="tabs-container">
        <div class="tabs">
            <div class="tab">🏠 Site Permissions</div>
            <div class="tab active">📚 Lists & Libraries</div>
        </div>
    </div>
    
    <div class="toolbar">
        <input type="text" class="search-box" placeholder="🔍 Search lists...">
        <button class="export-btn">📥 Export</button>
    </div>
    
    <div class="content">
        <!-- List Card 1: Unique Permissions -->
        <div class="list-card">
            <div class="list-card-header">
                <div class="list-title">
                    <span class="list-icon">📁</span>
                    <h3>Project Documents</h3>
                </div>
                <span class="permission-badge badge-unique">🔒 UNIQUE</span>
            </div>
            <div class="list-url">/sites/mysite/ProjectDocuments</div>
            
            <table class="permission-table">
                <thead>
                    <tr>
                        <th>User/Group</th>
                        <th>Type</th>
                        <th>Permission Level</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>
                            <div class="user-cell">
                                <div class="avatar avatar-blue">JS</div>
                                <div class="user-info">
                                    <span class="user-name">Jane Smith</span>
                                    <span class="user-email">jane.smith@company.com</span>
                                </div>
                            </div>
                        </td>
                        <td>User</td>
                        <td><span class="badge badge-full-control">Full Control</span></td>
                    </tr>
                    <tr>
                        <td>
                            <div class="user-cell">
                                <div class="avatar avatar-purple">PT</div>
                                <div class="user-info">
                                    <span class="user-name">Project Team</span>
                                    <span class="user-email">25 members</span>
                                </div>
                            </div>
                        </td>
                        <td>SharePoint Group</td>
                        <td><span class="badge badge-contribute">Contribute</span></td>
                    </tr>
                    <tr>
                        <td>
                            <div class="user-cell">
                                <div class="avatar avatar-orange">EP</div>
                                <div class="user-info">
                                    <span class="user-name">External Partners</span>
                                    <span class="user-email">external@partners.com</span>
                                </div>
                            </div>
                        </td>
                        <td>Security Group</td>
                        <td><span class="badge badge-read">Read</span></td>
                    </tr>
                </tbody>
            </table>
        </div>
        
        <!-- List Card 2: Inherited Permissions -->
        <div class="list-card">
            <div class="list-card-header">
                <div class="list-title">
                    <span class="list-icon">📋</span>
                    <h3>Customer List</h3>
                </div>
                <span class="permission-badge badge-inherited">⬆️ INHERITED</span>
            </div>
            <div class="list-url">/sites/mysite/Lists/Customers</div>
            
            <div class="inherited-info">
                <span class="inherited-info-icon">ℹ️</span>
                <div>
                    <div class="inherited-info-text">This list inherits permissions from the parent site</div>
                    <a href="#" style="color: #0078d4; font-size: 12px;">👁️ View Site Permissions</a>
                </div>
            </div>
        </div>
        
        <!-- List Card 3: Unique Permissions -->
        <div class="list-card">
            <div class="list-card-header">
                <div class="list-title">
                    <span class="list-icon">📁</span>
                    <h3>Shared Documents</h3>
                </div>
                <span class="permission-badge badge-unique">🔒 UNIQUE</span>
            </div>
            <div class="list-url">/sites/mysite/Shared Documents</div>
            
            <table class="permission-table">
                <thead>
                    <tr>
                        <th>User/Group</th>
                        <th>Type</th>
                        <th>Permission Level</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>
                            <div class="user-cell">
                                <div class="avatar avatar-green">EV</div>
                                <div class="user-info">
                                    <span class="user-name">Everyone</span>
                                    <span class="user-email">All users</span>
                                </div>
                            </div>
                        </td>
                        <td>Security Group</td>
                        <td><span class="badge badge-read">Read</span></td>
                    </tr>
                    <tr>
                        <td>
                            <div class="user-cell">
                                <div class="avatar avatar-blue">HR</div>
                                <div class="user-info">
                                    <span class="user-name">HR Department</span>
                                    <span class="user-email">hr@company.com</span>
                                </div>
                            </div>
                        </td>
                        <td>Security Group</td>
                        <td><span class="badge badge-contribute">Contribute</span></td>
                    </tr>
                    <tr>
                        <td>
                            <div class="user-cell">
                                <div class="avatar avatar-orange">MJ</div>
                                <div class="user-info">
                                    <span class="user-name">Mike Johnson</span>
                                    <span class="user-email">mike.johnson@company.com</span>
                                </div>
                            </div>
                        </td>
                        <td>User</td>
                        <td><span class="badge badge-edit">Edit</span></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</div>

<br><br>

<!-- LOADING STATE -->
<div class="webpart-container">
    <div class="header">
        <h1>📊 SharePoint Permission Viewer</h1>
        <button class="refresh-btn">🔄 Refresh</button>
    </div>
    
    <div class="loading-state">
        <div class="spinner"></div>
        <div class="loading-text">🔄 Loading site permissions...</div>
        <div class="progress-bar">
            <div class="progress-fill"></div>
        </div>
        <div style="font-size: 14px; color: #605e5c; margin-top: 16px;">
            This may take a moment for sites with many lists...
        </div>
    </div>
</div>

</body>
</html>