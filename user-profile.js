// Get user email from URL parameter
const urlParams = new URLSearchParams(window.location.search);
const userEmail = urlParams.get('email');

// Mock Data (same as main app)
const mockData = {
    users: [
        { name: 'Al Manji | Quest AV', email: 'al@quest-av.com', license: 'Office 365 E5', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Executive', created: '2020-01-15', title: 'President', phone: '(416) 340-7368 x240' },
        { name: 'Robert Styka | Quest AV', email: 'robert@quest-av.com', license: 'Office 365 E5', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Operations', created: '2020-02-01', title: 'General Manager', phone: '(416) 340-7368 x242' },
        { name: 'Natasha Manji | Quest AV', email: 'natasha@quest-av.com', license: 'Office 365 E5', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Sales', created: '2020-03-15', title: 'Director, Sales', phone: '(416) 340-7368 x245' },
        { name: 'Whitney Wilson | Quest AV', email: 'whitney@quest-av.com', license: 'Office 365 E5', mfa: true, lastSignIn: '2025-11-15', status: 'Active', department: 'Production Services', created: '2020-04-20', title: 'Director, Production Services', phone: '(416) 340-7368 x219' },
        { name: 'Adam Boyle | Quest AV', email: 'adam@quest-av.com', license: 'Office 365 E5', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Operations', created: '2020-05-10', title: 'Director, Operations', phone: '(416) 340-7368 x297' },
        { name: 'Nicholas Persaud | Quest AV', email: 'nicholas@quest-av.com', license: 'Office 365 E5', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Technology', created: '2020-06-01', title: 'Chief Technology Officer (CTO)', phone: '(416) 340-7368 x231' },
        { name: 'Noah Styka | Quest AV', email: 'noah@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-15', status: 'Active', department: 'Production', created: '2022-03-15', title: 'Production Specialist' },
        { name: 'Sushan Dsouza | Quest AV', email: 'sushan@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Operations', created: '2022-07-20', title: 'Operations Coordinator' },
        { name: 'TJ Humble | Quest AV', email: 'tj@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-15', status: 'Active', department: 'Technical', created: '2023-01-10', title: 'Technical Specialist' },
        { name: 'Hussein Esmail | Quest AV', email: 'hussein@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Sales', created: '2023-05-15', title: 'Sales Coordinator' },
        { name: 'Alan McArdle | Quest AV', email: 'alan@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-15', status: 'Active', department: 'Sales', created: '2023-08-20', title: 'Sales Representative' },
        { name: 'Jason Allen | Quest AV', email: 'jason@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-16', status: 'Active', department: 'Operations', created: '2024-02-01', title: 'Operations Assistant' },
        { name: 'Alejandro Bravo | Quest AV', email: 'alejandro@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-15', status: 'Active', department: 'Production', created: '2022-09-01', title: 'Production Manager' },
        { name: 'Alfred Park | Quest AV', email: 'alfred@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-14', status: 'Active', department: 'Operations', created: '2022-10-15', title: 'Operations Manager' },
        { name: 'Mohammad Kakour | Quest AV', email: 'Mohammad@quest-av.com', license: 'Office 365 E3', mfa: true, lastSignIn: '2025-11-15', status: 'Active', department: 'Technical', created: '2023-02-20', title: 'Technical Manager' }
    ],
    
    permissions: [
        { user: 'Al Manji | Quest AV', role: 'Global Administrator', service: 'All Services', assignedDate: '2020-01-15', assignedBy: 'System Admin' },
        { user: 'Robert Styka | Quest AV', role: 'Global Administrator', service: 'All Services', assignedDate: '2020-02-01', assignedBy: 'Al Manji | Quest AV' },
        { user: 'Nicholas Persaud | Quest AV', role: 'Exchange Administrator', service: 'Exchange Online', assignedDate: '2020-06-01', assignedBy: 'Al Manji | Quest AV' },
        { user: 'Nicholas Persaud | Quest AV', role: 'SharePoint Administrator', service: 'SharePoint Online', assignedDate: '2020-06-01', assignedBy: 'Al Manji | Quest AV' },
        { user: 'Natasha Manji | Quest AV', role: 'User Administrator', service: 'Azure AD', assignedDate: '2020-03-15', assignedBy: 'Al Manji | Quest AV' },
        { user: 'Adam Boyle | Quest AV', role: 'Teams Administrator', service: 'Microsoft Teams', assignedDate: '2020-05-10', assignedBy: 'Robert Styka | Quest AV' }
    ],

    mailboxPermissions: [
        { mailbox: 'sales@quest-av.com', user: 'Natasha Manji | Quest AV', type: 'Full Access', grantedDate: '2020-04-01' },
        { mailbox: 'sales@quest-av.com', user: 'Hussein Esmail | Quest AV', type: 'Send As', grantedDate: '2023-06-01' },
        { mailbox: 'sales@quest-av.com', user: 'Alan McArdle | Quest AV', type: 'Send As', grantedDate: '2023-09-01' },
        { mailbox: 'info@quest-av.com', user: 'Robert Styka | Quest AV', type: 'Full Access', grantedDate: '2020-02-15' },
        { mailbox: 'production@quest-av.com', user: 'Whitney Wilson | Quest AV', type: 'Full Access', grantedDate: '2020-05-01' }
    ],

    groups: [
        { name: 'Executive Team', type: 'Microsoft 365', members: 6, owners: 1, created: '2020-01-20' },
        { name: 'Sales Team', type: 'Microsoft 365', members: 3, owners: 1, created: '2020-03-20' },
        { name: 'Production Team', type: 'Microsoft 365', members: 5, owners: 1, created: '2020-04-25' },
        { name: 'Operations Team', type: 'Security', members: 8, owners: 2, created: '2020-02-10' },
        { name: 'Technical Team', type: 'Security', members: 4, owners: 1, created: '2020-06-15' },
        { name: 'All Quest AV', type: 'Dynamic', members: 15, owners: 1, created: '2020-01-15' },
        { name: 'Management Team', type: 'Microsoft 365', members: 9, owners: 1, created: '2020-02-01' }
    ]
};

// Helper functions (same as main app)
function getEnabledServices(license) {
    const serviceMap = {
        'Office 365 E5': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Yammer', 'Power BI Pro', 'Power Apps', 'Power Automate', 'Stream', 'Forms', 'Planner', 'To-Do', 'Delve', 'Sway', 'Advanced Threat Protection', 'Azure Information Protection'],
        'Office 365 E3': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Yammer', 'Power Apps', 'Power Automate', 'Stream', 'Forms', 'Planner', 'To-Do', 'Delve', 'Sway'],
        'Office 365 E1': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Yammer', 'Forms', 'Planner', 'To-Do'],
        'Business Premium': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Outlook', 'Word', 'Excel', 'PowerPoint', 'OneNote']
    };
    return serviceMap[license] || [];
}

function getUserGroups(userName) {
    const userGroupMap = {
        'Al Manji | Quest AV': ['Executive Team', 'Management Team', 'All Quest AV'],
        'Robert Styka | Quest AV': ['Executive Team', 'Management Team', 'Operations Team', 'All Quest AV'],
        'Natasha Manji | Quest AV': ['Executive Team', 'Management Team', 'Sales Team', 'All Quest AV'],
        'Whitney Wilson | Quest AV': ['Executive Team', 'Management Team', 'Production Team', 'All Quest AV'],
        'Adam Boyle | Quest AV': ['Executive Team', 'Management Team', 'Operations Team', 'All Quest AV'],
        'Nicholas Persaud | Quest AV': ['Executive Team', 'Management Team', 'Technical Team', 'All Quest AV'],
        'Noah Styka | Quest AV': ['Production Team', 'All Quest AV'],
        'Sushan Dsouza | Quest AV': ['Operations Team', 'All Quest AV'],
        'TJ Humble | Quest AV': ['Technical Team', 'All Quest AV'],
        'Hussein Esmail | Quest AV': ['Sales Team', 'All Quest AV'],
        'Alan McArdle | Quest AV': ['Sales Team', 'All Quest AV'],
        'Jason Allen | Quest AV': ['Operations Team', 'All Quest AV'],
        'Alejandro Bravo | Quest AV': ['Production Team', 'Management Team', 'All Quest AV'],
        'Alfred Park | Quest AV': ['Operations Team', 'Management Team', 'All Quest AV'],
        'Mohammad Kakour | Quest AV': ['Technical Team', 'Management Team', 'All Quest AV']
    };
    
    const userGroups = userGroupMap[userName] || [];
    return mockData.groups.filter(g => userGroups.includes(g.name)).map(g => ({
        ...g,
        role: Math.random() > 0.7 ? 'Owner' : 'Member'
    }));
}

function getUserMailboxData(user) {
    return {
        size: (Math.random() * 40 + 10).toFixed(2) + ' GB',
        itemCount: Math.floor(Math.random() * 50000 + 10000),
        quotaLimit: user.license.includes('E5') ? '100 GB' : user.license.includes('E3') ? '50 GB' : '25 GB',
        aliases: [user.email, user.email.replace('@', '.alt@')],
        permissions: mockData.mailboxPermissions.filter(p => p.user === user.name),
        archiveEnabled: user.license.includes('E5') || user.license.includes('E3'),
        litigationHold: user.department === 'Executive' || user.department === 'Finance'
    };
}

function getUserSharePointData(user) {
    return {
        storage: {
            used: (Math.random() * 800 + 100).toFixed(2) + ' GB',
            quota: '1 TB',
            percentage: Math.floor(Math.random() * 60 + 20)
        },
        sites: [
            { name: user.department + ' Team Site', role: 'Member', url: '/sites/' + user.department.toLowerCase().replace(/\s/g, '') },
            { name: 'Quest AV Intranet', role: 'Visitor', url: '/sites/intranet' },
            { name: 'Projects Hub', role: 'Member', url: '/sites/projects' }
        ],
        sharedFiles: Math.floor(Math.random() * 150 + 50),
        externalSharing: user.license.includes('E5')
    };
}

function getUserTeamsData(user) {
    const teamMap = {
        'Executive': ['Executive Team', 'All Quest AV', 'Management Team'],
        'Sales': ['Sales Team', 'All Quest AV'],
        'Production Services': ['Production Team', 'All Quest AV'],
        'Production': ['Production Team', 'All Quest AV'],
        'Operations': ['Operations Team', 'All Quest AV', 'Management Team'],
        'Technology': ['Technical Team', 'All Quest AV', 'Management Team'],
        'Technical': ['Technical Team', 'All Quest AV']
    };
    
    return {
        teams: (teamMap[user.department] || ['All Quest AV']).map(name => ({
            name: name,
            role: Math.random() > 0.8 ? 'Owner' : 'Member',
            channels: Math.floor(Math.random() * 10 + 3)
        })),
        guestAccessEnabled: true,
        externalAccess: user.license.includes('E5')
    };
}

function getUserSecurityData(user) {
    return {
        mfaStatus: user.mfa,
        mfaMethod: user.mfa ? 'Microsoft Authenticator App' : 'Not Configured',
        riskLevel: user.mfa ? 'Low' : 'Medium',
        conditionalAccessPolicies: ['Require MFA for All Users', 'Block Legacy Authentication', 'Require Compliant Device'],
        compliancePolicies: ['Data Loss Prevention', 'Retention Policy - 7 Years'],
        lastPasswordChange: getRandomPastDate(90),
        suspiciousActivity: false,
        blockedSignIns: Math.floor(Math.random() * 5)
    };
}

function getUserActivityData(user) {
    return {
        signIns: generateSignInData(),
        recentActivities: [
            { action: 'Read email', service: 'Exchange', time: '2 hours ago', icon: 'envelope' },
            { action: 'Edited document', service: 'SharePoint', time: '4 hours ago', icon: 'file-alt' },
            { action: 'Teams meeting', service: 'Teams', time: '6 hours ago', icon: 'video' },
            { action: 'Created file', service: 'OneDrive', time: '1 day ago', icon: 'folder' },
            { action: 'Signed in', service: 'Azure AD', time: '1 day ago', icon: 'sign-in-alt' }
        ],
        serviceUsage: {
            Exchange: Math.floor(Math.random() * 40 + 60),
            Teams: Math.floor(Math.random() * 30 + 50),
            SharePoint: Math.floor(Math.random() * 20 + 40),
            OneDrive: Math.floor(Math.random() * 25 + 45)
        }
    };
}

function getExternalFilesData(user) {
    const fileTypes = ['Budget_2025.xlsx', 'Project_Proposal.docx', 'Marketing_Plan.pptx', 'Client_Data.xlsx', 'Strategy_Document.pdf'];
    const externalUsers = ['john@partner.com', 'maria@client.com', 'tom@vendor.com', 'sarah@consultant.com'];
    
    return Array.from({ length: Math.floor(Math.random() * 8 + 3) }, (_, i) => ({
        fileName: fileTypes[Math.floor(Math.random() * fileTypes.length)].replace('.', `_${i+1}.`),
        sharedWith: externalUsers[Math.floor(Math.random() * externalUsers.length)],
        sharedDate: getRandomPastDate(60),
        accessType: Math.random() > 0.5 ? 'View Only' : 'Edit',
        location: Math.random() > 0.5 ? 'OneDrive' : 'SharePoint'
    }));
}

function getExternalLinksData(user) {
    const linkTypes = ['Quarterly Report', 'Project Files', 'Team Folder', 'Client Presentation', 'Budget Spreadsheet'];
    
    return Array.from({ length: Math.floor(Math.random() * 5 + 2) }, (_, i) => ({
        linkName: linkTypes[Math.floor(Math.random() * linkTypes.length)] + ` - Link ${i+1}`,
        url: `https://company.sharepoint.com/shared/${Math.random().toString(36).substring(7)}`,
        createdDate: getRandomPastDate(90),
        expiryDate: Math.random() > 0.6 ? 'Never' : getRandomFutureDate(30),
        accessCount: Math.floor(Math.random() * 50 + 5),
        allowAnonymous: Math.random() > 0.5
    }));
}

function getRandomPastDate(daysAgo) {
    const date = new Date();
    date.setDate(date.getDate() - Math.floor(Math.random() * daysAgo));
    return date.toISOString().split('T')[0];
}

function getRandomFutureDate(daysAhead) {
    const date = new Date();
    date.setDate(date.getDate() + Math.floor(Math.random() * daysAhead + 7));
    return date.toISOString().split('T')[0];
}

function generateSignInData() {
    return Array.from({length: 30}, (_, i) => ({
        day: i + 1,
        count: Math.floor(Math.random() * 15 + 5)
    }));
}

function getUserProfileData(email) {
    const user = mockData.users.find(u => u.email === email);
    if (!user) return null;

    return {
        basic: user,
        licenses: {
            primary: user.license,
            services: getEnabledServices(user.license)
        },
        adminRoles: mockData.permissions.filter(p => p.user === user.name),
        groups: getUserGroups(user.name),
        mailbox: getUserMailboxData(user),
        sharepoint: getUserSharePointData(user),
        teams: getUserTeamsData(user),
        security: getUserSecurityData(user),
        activity: getUserActivityData(user),
        externalFiles: getExternalFilesData(user),
        externalLinks: getExternalLinksData(user)
    };
}

// Initialize page on load
document.addEventListener('DOMContentLoaded', function() {
    if (!userEmail) {
        alert('No user specified');
        window.location.href = 'index.html';
        return;
    }

    const profileData = getUserProfileData(userEmail);
    if (!profileData) {
        alert('User not found');
        window.location.href = 'index.html';
        return;
    }

    // Populate header
    document.getElementById('profileUserName').textContent = profileData.basic.name;
    document.getElementById('profileUserEmail').textContent = profileData.basic.email;
    document.getElementById('profileUserStatus').textContent = profileData.basic.status;
    document.getElementById('profileUserStatus').className = 'badge ' + (profileData.basic.status === 'Active' ? 'badge-success' : 'badge-secondary');
    document.getElementById('profileUserMFA').textContent = profileData.basic.mfa ? 'MFA Enabled' : 'MFA Disabled';
    document.getElementById('profileUserMFA').className = 'badge ' + (profileData.basic.mfa ? 'badge-success' : 'badge-warning');

    // Update page title
    document.title = `${profileData.basic.name} - User Profile`;

    // Initialize tabs
    initializeProfileTabs();

    // Populate all tabs
    populateOverviewTab(profileData);
    populateLicensesTab(profileData);
    populatePermissionsTab(profileData);
    populateGroupsTab(profileData);
    populateMailboxTab(profileData);
    populateSharePointTab(profileData);
    populateTeamsTab(profileData);
    populateExternalSharingTab(profileData);
    populateSecurityTab(profileData);
    populateActivityTab(profileData);
});

// Tab initialization
function initializeProfileTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            tabBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });

            const tabId = this.getAttribute('data-tab') + '-tab';
            document.getElementById(tabId).classList.add('active');
        });
    });
}

// Populate functions
function populateOverviewTab(data) {
    document.getElementById('profileLastSignIn').textContent = data.basic.lastSignIn;
    document.getElementById('profileGroupCount').textContent = data.groups.length + ' groups';
    document.getElementById('profileAdminRoles').textContent = data.adminRoles.length > 0 ? data.adminRoles.length + ' roles' : 'No admin roles';
    document.getElementById('profileStorage').textContent = data.sharepoint.storage.used;
    document.getElementById('profileDepartment').textContent = data.basic.department;
    document.getElementById('profileCreated').textContent = data.basic.created;
    document.getElementById('profileAccountStatus').textContent = data.basic.status;
    document.getElementById('profileLicenseType').textContent = data.basic.license;
}

function populateLicensesTab(data) {
    const licensesDiv = document.getElementById('profileLicenses');
    licensesDiv.innerHTML = `
        <div class="service-item">
            <div>
                <h4>${data.licenses.primary}</h4>
                <p>Primary license assigned</p>
            </div>
            <span class="badge badge-success">Active</span>
        </div>
    `;

    const servicesDiv = document.getElementById('profileServices');
    servicesDiv.innerHTML = data.licenses.services.map(service => `
        <div class="service-card enabled">
            <i class="fas fa-check-circle"></i>
            <p>${service}</p>
        </div>
    `).join('');
}

function populatePermissionsTab(data) {
    const adminRolesDiv = document.getElementById('profileAdminRolesList');
    
    if (data.adminRoles.length > 0) {
        adminRolesDiv.innerHTML = data.adminRoles.map(role => `
            <div class="permission-item">
                <div>
                    <h4>${role.role}</h4>
                    <p>${role.service} - Assigned on ${role.assignedDate} by ${role.assignedBy}</p>
                </div>
                <span class="badge badge-danger">Admin</span>
            </div>
        `).join('');
    } else {
        adminRolesDiv.innerHTML = '<p style="color: var(--text-secondary); padding: 1rem;">No administrative roles assigned</p>';
    }

    const delegatedDiv = document.getElementById('profileDelegatedPerms');
    if (data.mailbox.permissions.length > 0) {
        delegatedDiv.innerHTML = data.mailbox.permissions.map(perm => `
            <div class="permission-item">
                <div>
                    <h4>${perm.type} - ${perm.mailbox}</h4>
                    <p>Granted on ${perm.grantedDate}</p>
                </div>
                <span class="badge badge-info">Mailbox</span>
            </div>
        `).join('');
    } else {
        delegatedDiv.innerHTML = '<p style="color: var(--text-secondary); padding: 1rem;">No delegated permissions</p>';
    }
}

function populateGroupsTab(data) {
    const groupsDiv = document.getElementById('profileGroupsList');
    const memberGroups = data.groups.filter(g => g.role === 'Member');
    const ownedGroups = data.groups.filter(g => g.role === 'Owner');

    groupsDiv.innerHTML = memberGroups.map(group => `
        <div class="group-item">
            <div>
                <h4>${group.name}</h4>
            </div>
            <div class="group-meta">
                <span class="badge badge-info">${group.type}</span>
                <span class="badge badge-secondary">${group.members} members</span>
            </div>
        </div>
    `).join('');

    const ownedDiv = document.getElementById('profileOwnedGroups');
    if (ownedGroups.length > 0) {
        ownedDiv.innerHTML = ownedGroups.map(group => `
            <div class="group-item">
                <div>
                    <h4>${group.name}</h4>
                </div>
                <div class="group-meta">
                    <span class="badge badge-info">${group.type}</span>
                    <span class="badge badge-warning">Owner</span>
                </div>
            </div>
        `).join('');
    } else {
        ownedDiv.innerHTML = '<p style="color: var(--text-secondary); padding: 1rem;">Not an owner of any groups</p>';
    }
}

function populateMailboxTab(data) {
    const mailboxInfo = document.getElementById('profileMailboxInfo');
    mailboxInfo.innerHTML = `
        <div class="info-item">
            <strong>Mailbox Size</strong>
            <span>${data.mailbox.size}</span>
        </div>
        <div class="info-item">
            <strong>Item Count</strong>
            <span>${data.mailbox.itemCount.toLocaleString()}</span>
        </div>
        <div class="info-item">
            <strong>Quota Limit</strong>
            <span>${data.mailbox.quotaLimit}</span>
        </div>
        <div class="info-item">
            <strong>Archive</strong>
            <span>${data.mailbox.archiveEnabled ? 'Enabled' : 'Disabled'}</span>
        </div>
        <div class="info-item">
            <strong>Litigation Hold</strong>
            <span>${data.mailbox.litigationHold ? 'Enabled' : 'Disabled'}</span>
        </div>
    `;

    const aliasesDiv = document.getElementById('profileEmailAliases');
    aliasesDiv.innerHTML = data.mailbox.aliases.map(alias => `
        <div class="service-item">
            <div>
                <h4>${alias}</h4>
                <p>Email alias</p>
            </div>
            <span class="badge badge-info">Active</span>
        </div>
    `).join('');

    const permsDiv = document.getElementById('profileMailboxPerms');
    if (data.mailbox.permissions.length > 0) {
        permsDiv.innerHTML = data.mailbox.permissions.map(perm => `
            <div class="permission-item">
                <div>
                    <h4>${perm.type}</h4>
                    <p>Access to ${perm.mailbox}</p>
                </div>
                <span class="badge badge-warning">Delegated</span>
            </div>
        `).join('');
    } else {
        permsDiv.innerHTML = '<p style="color: var(--text-secondary); padding: 1rem;">No mailbox permissions granted</p>';
    }
}

function populateSharePointTab(data) {
    const oneDriveInfo = document.getElementById('profileOneDriveInfo');
    oneDriveInfo.innerHTML = `
        <div class="info-item">
            <strong>Storage Used</strong>
            <span>${data.sharepoint.storage.used}</span>
        </div>
        <div class="info-item">
            <strong>Storage Quota</strong>
            <span>${data.sharepoint.storage.quota}</span>
        </div>
        <div class="info-item">
            <strong>Usage</strong>
            <span>${data.sharepoint.storage.percentage}%</span>
        </div>
        <div class="info-item">
            <strong>Files Shared</strong>
            <span>${data.sharepoint.sharedFiles}</span>
        </div>
    `;

    const sitesDiv = document.getElementById('profileSharePointSites');
    sitesDiv.innerHTML = data.sharepoint.sites.map(site => `
        <div class="group-item">
            <div>
                <h4>${site.name}</h4>
                <p style="color: var(--text-secondary); font-size: 0.85rem;">${site.url}</p>
            </div>
            <span class="badge badge-info">${site.role}</span>
        </div>
    `).join('');

    const filesDiv = document.getElementById('profileSharedFiles');
    filesDiv.innerHTML = `
        <p style="color: var(--text-primary); padding: 1rem;">
            <strong>${data.sharepoint.sharedFiles}</strong> files currently shared with external users
        </p>
    `;
}

function populateTeamsTab(data) {
    const teamsDiv = document.getElementById('profileTeamsList');
    teamsDiv.innerHTML = data.teams.teams.map(team => `
        <div class="group-item">
            <div>
                <h4>${team.name}</h4>
                <p style="color: var(--text-secondary); font-size: 0.85rem;">${team.channels} channels</p>
            </div>
            <div class="group-meta">
                <span class="badge badge-info">${team.role}</span>
            </div>
        </div>
    `).join('');

    const settingsDiv = document.getElementById('profileTeamsSettings');
    settingsDiv.innerHTML = `
        <div class="info-item">
            <strong>Guest Access</strong>
            <span>${data.teams.guestAccessEnabled ? 'Enabled' : 'Disabled'}</span>
        </div>
        <div class="info-item">
            <strong>External Access</strong>
            <span>${data.teams.externalAccess ? 'Enabled' : 'Disabled'}</span>
        </div>
        <div class="info-item">
            <strong>Total Teams</strong>
            <span>${data.teams.teams.length}</span>
        </div>
    `;
}

function populateExternalSharingTab(data) {
    // External Files
    const externalFilesDiv = document.getElementById('profileExternalFiles');
    if (data.externalFiles && data.externalFiles.length > 0) {
        externalFilesDiv.innerHTML = data.externalFiles.map(file => `
            <div class="external-item">
                <div class="external-item-info">
                    <h4><i class="fas fa-file"></i> ${file.fileName}</h4>
                    <p>Shared with <strong>${file.sharedWith}</strong> on ${file.sharedDate}</p>
                </div>
                <div class="external-meta">
                    <span class="badge badge-warning">${file.accessType}</span>
                    <span class="badge badge-info">${file.location}</span>
                </div>
            </div>
        `).join('');
    } else {
        externalFilesDiv.innerHTML = '<p style="color: var(--text-secondary); padding: 1rem;">No files shared externally</p>';
    }

    // External Links
    const externalLinksDiv = document.getElementById('profileExternalLinks');
    if (data.externalLinks && data.externalLinks.length > 0) {
        externalLinksDiv.innerHTML = data.externalLinks.map(link => `
            <div class="link-item">
                <div class="link-icon">
                    <i class="fas fa-link"></i>
                </div>
                <div class="link-info">
                    <h4>${link.linkName}</h4>
                    <p>Created: ${link.createdDate} | Expires: ${link.expiryDate} | Access count: ${link.accessCount}</p>
                    <p class="link-url">${link.url}</p>
                </div>
                <span class="badge ${link.allowAnonymous ? 'badge-warning' : 'badge-info'}">
                    ${link.allowAnonymous ? 'Anonymous Access' : 'Authenticated Only'}
                </span>
            </div>
        `).join('');
    } else {
        externalLinksDiv.innerHTML = '<p style="color: var(--text-secondary); padding: 1rem;">No external links created</p>';
    }

    // Sharing Settings
    const settingsDiv = document.getElementById('profileSharingSettings');
    settingsDiv.innerHTML = `
        <div class="info-item">
            <strong>External Sharing</strong>
            <span>${data.sharepoint.externalSharing ? 'Enabled' : 'Disabled'}</span>
        </div>
        <div class="info-item">
            <strong>Total External Files</strong>
            <span>${data.externalFiles.length}</span>
        </div>
        <div class="info-item">
            <strong>Total External Links</strong>
            <span>${data.externalLinks.length}</span>
        </div>
    `;
}

function populateSecurityTab(data) {
    const securityDiv = document.getElementById('profileSecurityStatus');
    securityDiv.innerHTML = `
        <div class="info-item">
            <strong>MFA Status</strong>
            <span style="color: ${data.security.mfaStatus ? 'var(--success-color)' : 'var(--danger-color)'}">
                ${data.security.mfaStatus ? 'Enabled' : 'Disabled'}
            </span>
        </div>
        <div class="info-item">
            <strong>MFA Method</strong>
            <span>${data.security.mfaMethod}</span>
        </div>
        <div class="info-item">
            <strong>Risk Level</strong>
            <span style="color: ${data.security.riskLevel === 'Low' ? 'var(--success-color)' : 'var(--warning-color)'}">
                ${data.security.riskLevel}
            </span>
        </div>
        <div class="info-item">
            <strong>Last Password Change</strong>
            <span>${data.security.lastPasswordChange}</span>
        </div>
        <div class="info-item">
            <strong>Blocked Sign-ins</strong>
            <span>${data.security.blockedSignIns}</span>
        </div>
        <div class="info-item">
            <strong>Suspicious Activity</strong>
            <span style="color: ${data.security.suspiciousActivity ? 'var(--danger-color)' : 'var(--success-color)'}">
                ${data.security.suspiciousActivity ? 'Detected' : 'None'}
            </span>
        </div>
    `;

    const caDiv = document.getElementById('profileConditionalAccess');
    caDiv.innerHTML = data.security.conditionalAccessPolicies.map(policy => `
        <div class="service-item">
            <div>
                <h4>${policy}</h4>
                <p>Applied to this user</p>
            </div>
            <span class="badge badge-success">Active</span>
        </div>
    `).join('');

    const complianceDiv = document.getElementById('profileCompliance');
    complianceDiv.innerHTML = data.security.compliancePolicies.map(policy => `
        <div class="service-item">
            <div>
                <h4>${policy}</h4>
                <p>Compliance policy applied</p>
            </div>
            <span class="badge badge-info">Applied</span>
        </div>
    `).join('');
}

function populateActivityTab(data) {
    // Sign-in activity chart
    const ctx = document.getElementById('signInActivityChart');
    if (ctx) {
        new Chart(ctx.getContext('2d'), {
            type: 'bar',
            data: {
                labels: data.activity.signIns.map(d => 'Day ' + d.day),
                datasets: [{
                    label: 'Sign-ins',
                    data: data.activity.signIns.map(d => d.count),
                    backgroundColor: 'rgba(160, 204, 58, 0.7)',
                    borderColor: '#A0CC3A',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
    }

    // Recent activities
    const activitiesDiv = document.getElementById('profileRecentActivities');
    activitiesDiv.innerHTML = data.activity.recentActivities.map(activity => `
        <div class="activity-item">
            <div class="activity-icon">
                <i class="fas fa-${activity.icon}"></i>
            </div>
            <div class="activity-info">
                <h4>${activity.action}</h4>
                <p>${activity.service}</p>
            </div>
            <div class="activity-time">${activity.time}</div>
        </div>
    `).join('');

    // Service usage
    const usageDiv = document.getElementById('profileServiceUsage');
    usageDiv.innerHTML = Object.entries(data.activity.serviceUsage).map(([service, percentage]) => `
        <div class="usage-item">
            <h4>${service}</h4>
            <div class="usage-bar">
                <div class="usage-fill" style="width: ${percentage}%; background: var(--quest-green);"></div>
            </div>
            <p>${percentage}%</p>
        </div>
    `).join('');
}

function exportUserProfile() {
    alert('User Profile Export\n\nThis would generate a comprehensive PDF or Excel report containing:\n\n• User information\n• License details\n• All permissions and roles\n• Group memberships\n• Mailbox statistics\n• External files and links\n• Activity logs\n• Security status\n• Compliance information');
}
