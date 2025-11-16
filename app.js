// Mock Data
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
    ],

    securityPolicies: [
        { name: 'Require MFA for All Users', status: 'Enabled', appliedTo: 'All Users', conditions: 'All cloud apps', lastModified: '2024-09-15' },
        { name: 'Block Legacy Authentication', status: 'Enabled', appliedTo: 'All Users', conditions: 'Exchange ActiveSync', lastModified: '2024-08-20' },
        { name: 'Require Compliant Device', status: 'Enabled', appliedTo: 'All Users', conditions: 'Office 365', lastModified: '2024-10-01' },
        { name: 'Block Sign-in Risk', status: 'Report-only', appliedTo: 'All Users', conditions: 'High risk sign-ins', lastModified: '2024-11-05' }
    ],

    guestUsers: [
        { name: 'Partner Consultant', email: 'partner@external.com', invitedBy: 'Robert Styka | Quest AV', accessLevel: 'Guest', invitedDate: '2024-09-10', status: 'Accepted' },
        { name: 'Tech Vendor', email: 'vendor@tech.com', invitedBy: 'Nicholas Persaud | Quest AV', accessLevel: 'Guest', invitedDate: '2024-10-15', status: 'Accepted' },
        { name: 'Client Representative', email: 'client@company.com', invitedBy: 'Natasha Manji | Quest AV', accessLevel: 'Guest', invitedDate: '2024-11-01', status: 'Pending' }
    ],

    applications: [
        { name: 'Salesforce', publisher: 'Salesforce.com', permissions: 'Read/Write User Data', users: 5, lastUsed: '2025-11-16', status: 'Active' },
        { name: 'Zoom', publisher: 'Zoom Video Communications', permissions: 'Calendar Access', users: 15, lastUsed: '2025-11-16', status: 'Active' },
        { name: 'DocuSign', publisher: 'DocuSign Inc.', permissions: 'Email, Profile', users: 8, lastUsed: '2025-11-15', status: 'Active' },
        { name: 'Power BI', publisher: 'Microsoft', permissions: 'Full Access', users: 12, lastUsed: '2025-11-16', status: 'Active' },
        { name: 'Adobe Creative Cloud', publisher: 'Adobe', permissions: 'User Profile', users: 6, lastUsed: '2025-11-14', status: 'Active' }
    ],

    recentActivities: [
        { date: '2025-11-16 14:30', user: 'Robert Styka | Quest AV', action: 'Assigned Role', resource: 'Teams Admin to Adam Boyle', status: 'Success' },
        { date: '2025-11-16 12:15', user: 'Natasha Manji | Quest AV', action: 'Modified Permission', resource: 'Mailbox: sales@quest-av.com', status: 'Success' },
        { date: '2025-11-16 10:45', user: 'Nicholas Persaud | Quest AV', action: 'Modified Permission', resource: 'SharePoint Site Access', status: 'Success' },
        { date: '2025-11-15 16:20', user: 'Adam Boyle | Quest AV', action: 'Created Group', resource: 'Operations Team Q4', status: 'Success' },
        { date: '2025-11-15 14:00', user: 'System', action: 'Policy Applied', resource: 'MFA Enforcement', status: 'Warning' }
    ]
};

// Initialize Dashboard
document.addEventListener('DOMContentLoaded', function() {
    initializeNavigation();
    loadOverviewData();
    loadUsersTable();
    loadPermissionsTable();
    loadGuestUsers();
    loadRecentActivities();
    loadOrgChart();
    initializeCharts();
    initializeFilters();
    initializeOrgChartFilters();
});

// Navigation
function initializeNavigation() {
    const navItems = document.querySelectorAll('.nav-item');
    
    navItems.forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            
            // Remove active class from all nav items
            navItems.forEach(nav => nav.classList.remove('active'));
            
            // Add active class to clicked item
            this.classList.add('active');
            
            // Hide all sections
            document.querySelectorAll('.content-section').forEach(section => {
                section.classList.remove('active');
            });
            
            // Show selected section
            const sectionId = this.getAttribute('data-section');
            document.getElementById(sectionId).classList.add('active');
        });
    });
}

// Load Overview Data
function loadOverviewData() {
    const totalUsersEl = document.getElementById('totalUsers');
    const activeUsersEl = document.getElementById('activeUsers');
    const totalLicensesEl = document.getElementById('totalLicenses');
    const adminCountEl = document.getElementById('adminCount');
    
    if (totalUsersEl) totalUsersEl.textContent = mockData.users.length;
    if (activeUsersEl) activeUsersEl.textContent = mockData.users.filter(u => u.status === 'Active').length;
    if (totalLicensesEl) totalLicensesEl.textContent = mockData.users.length;
    if (adminCountEl) adminCountEl.textContent = new Set(mockData.permissions.map(p => p.user)).size;
}

// Load Users Table
function loadUsersTable() {
    const tbody = document.getElementById('usersTable');
    if (!tbody) return;
    
    tbody.innerHTML = mockData.users.map(user => `
        <tr>
            <td><strong>${user.name}</strong></td>
            <td>${user.email}</td>
            <td><span class="badge badge-info">${user.license}</span></td>
            <td><span class="badge ${user.mfa ? 'badge-success' : 'badge-warning'}">${user.mfa ? 'Enabled' : 'Disabled'}</span></td>
            <td>${user.lastSignIn}</td>
            <td><span class="badge ${user.status === 'Active' ? 'badge-success' : 'badge-secondary'}">${user.status}</span></td>
            <td>
                <div class="action-buttons">
                    <button class="action-btn action-btn-primary" onclick="viewUserDetails('${user.email}')" title="View Details">
                        <i class="fas fa-eye"></i>
                    </button>
                    <button class="action-btn action-btn-secondary" onclick="showClonePermissionsModal('${user.email}')" title="Clone Permissions">
                        <i class="fas fa-clone"></i>
                    </button>
                    <button class="action-btn action-btn-success" onclick="showQuickActionsModal('${user.email}')" title="Quick Actions">
                        <i class="fas fa-bolt"></i>
                    </button>
                    <button class="action-btn action-btn-warning" onclick="showAssignRoleModal('${user.email}')" title="Assign Role">
                        <i class="fas fa-user-shield"></i>
                    </button>
                </div>
            </td>
        </tr>
    `).join('');
}

// Load Permissions Table
function loadPermissionsTable() {
    const tbody = document.getElementById('permissionsTable');
    if (!tbody) return;
    
    tbody.innerHTML = mockData.permissions.map(perm => `
        <tr>
            <td><strong>${perm.user}</strong></td>
            <td><span class="badge badge-danger">${perm.role}</span></td>
            <td>${perm.service}</td>
            <td>${perm.assignedDate}</td>
            <td>${perm.assignedBy}</td>
            <td>
                <button class="action-btn">
                    <i class="fas fa-edit"></i> Modify
                </button>
            </td>
        </tr>
    `).join('');

    const mailboxTbody = document.getElementById('mailboxPermissionsTable');
    if (mailboxTbody) {
        mailboxTbody.innerHTML = mockData.mailboxPermissions.map(perm => `
            <tr>
                <td><strong>${perm.mailbox}</strong></td>
                <td>${perm.user}</td>
                <td><span class="badge badge-info">${perm.type}</span></td>
                <td>${perm.grantedDate}</td>
            </tr>
        `).join('');
    }
}

// Load Groups Table
function loadGroupsTable() {
    const tbody = document.getElementById('groupsTable');
    tbody.innerHTML = mockData.groups.map(group => `
        <tr>
            <td><strong>${group.name}</strong></td>
            <td><span class="badge badge-info">${group.type}</span></td>
            <td>${group.members}</td>
            <td>${group.owners}</td>
            <td>${group.created}</td>
            <td>
                <button class="action-btn">
                    <i class="fas fa-users"></i> Members
                </button>
            </td>
        </tr>
    `).join('');
}

// Load Security Policies
function loadSecurityPolicies() {
    const tbody = document.getElementById('securityPoliciesTable');
    tbody.innerHTML = mockData.securityPolicies.map(policy => `
        <tr>
            <td><strong>${policy.name}</strong></td>
            <td><span class="badge ${policy.status === 'Enabled' ? 'badge-success' : 'badge-warning'}">${policy.status}</span></td>
            <td>${policy.appliedTo}</td>
            <td>${policy.conditions}</td>
            <td>${policy.lastModified}</td>
        </tr>
    `).join('');
}

// Load Guest Users
function loadGuestUsers() {
    const tbody = document.getElementById('guestUsersTable');
    if (!tbody) return;
    
    tbody.innerHTML = mockData.guestUsers.map(guest => `
        <tr>
            <td><strong>${guest.name}</strong></td>
            <td>${guest.email}</td>
            <td>${guest.invitedBy}</td>
            <td>${guest.accessLevel}</td>
            <td>${guest.invitedDate}</td>
            <td><span class="badge ${guest.status === 'Accepted' ? 'badge-success' : 'badge-warning'}">${guest.status}</span></td>
        </tr>
    `).join('');
}

// Load Applications
function loadApplications() {
    const tbody = document.getElementById('appsTable');
    tbody.innerHTML = mockData.applications.map(app => `
        <tr>
            <td><strong>${app.name}</strong></td>
            <td>${app.publisher}</td>
            <td>${app.permissions}</td>
            <td>${app.users}</td>
            <td>${app.lastUsed}</td>
            <td><span class="badge badge-success">${app.status}</span></td>
        </tr>
    `).join('');
}

// Load Recent Activities
function loadRecentActivities() {
    const tbody = document.getElementById('recentActivities');
    if (!tbody) return;
    
    tbody.innerHTML = mockData.recentActivities.map(activity => `
        <tr>
            <td>${activity.date}</td>
            <td><strong>${activity.user}</strong></td>
            <td>${activity.action}</td>
            <td>${activity.resource}</td>
            <td><span class="badge ${activity.status === 'Success' ? 'badge-success' : 'badge-warning'}">${activity.status}</span></td>
        </tr>
    `).join('');
}

// Load Organization Chart
function loadOrgChart() {
    const container = document.getElementById('orgChartContainer');
    if (!container) return;

    const orgStructure = {
        ceo: { name: 'Al Manji | Quest AV', title: 'President', department: 'Executive', email: 'al@quest-av.com', reports: 6, phone: '(416) 340-7368 x240' },
        executives: [
            { name: 'Robert Styka | Quest AV', title: 'General Manager', department: 'Operations', email: 'robert@quest-av.com', reports: 5, phone: '(416) 340-7368 x242' },
            { name: 'Natasha Manji | Quest AV', title: 'Director, Sales', department: 'Sales', email: 'natasha@quest-av.com', reports: 3, phone: '(416) 340-7368 x245' },
            { name: 'Whitney Wilson | Quest AV', title: 'Director, Production Services', department: 'Production Services', email: 'whitney@quest-av.com', reports: 4, phone: '(416) 340-7368 x219' },
            { name: 'Adam Boyle | Quest AV', title: 'Director, Operations', department: 'Operations', email: 'adam@quest-av.com', reports: 3, phone: '(416) 340-7368 x297' },
            { name: 'Nicholas Persaud | Quest AV', title: 'Chief Technology Officer', department: 'Technology', email: 'nicholas@quest-av.com', reports: 2, phone: '(416) 340-7368 x231' }
        ],
        managers: [
            { 
                name: 'Alejandro Bravo | Quest AV', 
                title: 'Production Manager', 
                department: 'Production', 
                parent: 'Whitney Wilson | Quest AV', 
                email: 'alejandro@quest-av.com',
                team: 2,
                members: ['Noah Styka | Quest AV']
            },
            { 
                name: 'Alfred Park | Quest AV', 
                title: 'Operations Manager', 
                department: 'Operations', 
                parent: 'Adam Boyle | Quest AV', 
                email: 'alfred@quest-av.com',
                team: 2,
                members: ['Sushan Dsouza | Quest AV', 'Jason Allen | Quest AV']
            },
            { 
                name: 'Mohammad Kakour | Quest AV', 
                title: 'Technical Manager', 
                department: 'Technical', 
                parent: 'Nicholas Persaud | Quest AV', 
                email: 'Mohammad@quest-av.com',
                team: 1,
                members: ['TJ Humble | Quest AV']
            }
        ],
        directReports: {
            'Natasha Manji | Quest AV': ['Hussein Esmail | Quest AV', 'Alan McArdle | Quest AV'],
            'Robert Styka | Quest AV': ['Adam Boyle | Quest AV']
        }
    };

    container.innerHTML = `
        <div class="org-chart">
            <!-- CEO Level -->
            <div class="org-level">
                <div class="org-level-title">Executive Leadership</div>
            </div>
            <div class="org-level">
                <div class="org-node executive" onclick="viewUserDetails('${orgStructure.ceo.email}')">
                    <div class="org-node-avatar">
                        <i class="fas fa-user-tie"></i>
                    </div>
                    <h4>${orgStructure.ceo.name}</h4>
                    <p>${orgStructure.ceo.title}</p>
                    <span class="org-node-department">${orgStructure.ceo.department}</span>
                    <div class="org-node-stats">
                        <div class="org-node-stat">
                            <strong>${orgStructure.ceo.reports}</strong>
                            <span>Direct Reports</span>
                        </div>
                    </div>
                </div>
            </div>

            <div class="org-connector"></div>

            <!-- Executive Level -->
            <div class="org-level">
                <div class="org-level-title">Department Heads</div>
            </div>
            <div class="org-level">
                ${orgStructure.executives.map(exec => `
                    <div class="org-node" onclick="viewUserDetails('${exec.email}')">
                        <div class="org-node-avatar">
                            <i class="fas fa-user-shield"></i>
                        </div>
                        <h4>${exec.name}</h4>
                        <p>${exec.title}</p>
                        <span class="org-node-department">${exec.department}</span>
                        <div class="org-node-stats">
                            <div class="org-node-stat">
                                <strong>${exec.reports}</strong>
                                <span>Team Size</span>
                            </div>
                        </div>
                    </div>
                `).join('')}
            </div>

            <div class="org-connector"></div>

            <!-- Manager Level -->
            <div class="org-level">
                <div class="org-level-title">Management Team</div>
            </div>
            <div class="org-level">
                ${orgStructure.managers.map(mgr => `
                    <div class="org-node" onclick="viewUserDetails('${mgr.email}')">
                        <div class="org-node-avatar">
                            <i class="fas fa-user"></i>
                        </div>
                        <h4>${mgr.name}</h4>
                        <p>${mgr.title}</p>
                        <span class="org-node-department">${mgr.department}</span>
                        <div class="org-node-stats">
                            <div class="org-node-stat">
                                <strong>${mgr.team}</strong>
                                <span>Team Members</span>
                            </div>
                        </div>
                        ${mgr.members && mgr.members.length > 0 ? `
                            <div class="org-node-team-members">
                                <div class="team-members-label">
                                    <i class="fas fa-users"></i> Direct Reports
                                </div>
                                ${mgr.members.map(member => `
                                    <div class="team-member-item">
                                        <i class="fas fa-user-circle"></i>
                                        <span>${member}</span>
                                    </div>
                                `).join('')}
                            </div>
                        ` : ''}
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

// Initialize Charts
function initializeCharts() {
    // License Distribution Chart
    const licenseCanvas = document.getElementById('licenseChart');
    if (licenseCanvas) {
        const licenseCtx = licenseCanvas.getContext('2d');
        const licenseCounts = {};
        mockData.users.forEach(user => {
            licenseCounts[user.license] = (licenseCounts[user.license] || 0) + 1;
        });

        new Chart(licenseCtx, {
            type: 'doughnut',
            data: {
                labels: Object.keys(licenseCounts),
                datasets: [{
                    data: Object.values(licenseCounts),
                    backgroundColor: [
                        '#A0CC3A',
                        '#231F20',
                        '#7AB82D',
                        '#6B6B6B',
                        '#8AB32F'
                    ],
                    borderWidth: 2,
                    borderColor: '#FFFFFF'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 10,
                            font: { family: 'Poppins', size: 11 }
                        }
                    }
                }
            }
        });
    }

    // Activity Chart
    const activityCanvas = document.getElementById('activityChart');
    if (activityCanvas) {
        const activityCtx = activityCanvas.getContext('2d');
        new Chart(activityCtx, {
            type: 'line',
            data: {
                labels: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
                datasets: [{
                    label: 'Active Users',
                    data: [235, 242, 238, 248, 230, 140, 120],
                    borderColor: '#A0CC3A',
                    backgroundColor: 'rgba(160, 204, 58, 0.1)',
                    tension: 0.4,
                    borderWidth: 2,
                    fill: true
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: false,
                        min: 100,
                        ticks: {
                            font: { family: 'Poppins', size: 10 }
                        }
                    },
                    x: {
                        ticks: {
                            font: { family: 'Poppins', size: 10 }
                        }
                    }
                }
            }
        });
    }
}

// Initialize Filters
function initializeFilters() {
    // User search filter
    const userSearch = document.getElementById('userSearch');
    if (userSearch) {
        userSearch.addEventListener('input', function() {
            const searchTerm = this.value.toLowerCase();
            filterUsersTable(searchTerm);
        });
    }

    // License filter
    const licenseFilter = document.getElementById('licenseFilter');
    if (licenseFilter) {
        licenseFilter.addEventListener('change', function() {
            const selectedLicense = this.value;
            filterUsersByLicense(selectedLicense);
        });
    }

    // Global search
    const globalSearch = document.getElementById('globalSearch');
    if (globalSearch) {
        globalSearch.addEventListener('input', function() {
            console.log('Global search:', this.value);
            // Implement global search across all data
        });
    }
}

// Filter Users Table
function filterUsersTable(searchTerm) {
    const filteredUsers = mockData.users.filter(user => 
        user.name.toLowerCase().includes(searchTerm) ||
        user.email.toLowerCase().includes(searchTerm) ||
        user.department.toLowerCase().includes(searchTerm)
    );
    
    const tbody = document.getElementById('usersTable');
    tbody.innerHTML = filteredUsers.map(user => `
        <tr>
            <td><strong>${user.name}</strong></td>
            <td>${user.email}</td>
            <td><span class="badge badge-info">${user.license}</span></td>
            <td><span class="badge ${user.mfa ? 'badge-success' : 'badge-warning'}">${user.mfa ? 'Enabled' : 'Disabled'}</span></td>
            <td>${user.lastSignIn}</td>
            <td><span class="badge ${user.status === 'Active' ? 'badge-success' : 'badge-secondary'}">${user.status}</span></td>
            <td>
                <button class="action-btn" onclick="viewUserDetails('${user.email}')">
                    <i class="fas fa-eye"></i> View
                </button>
            </td>
        </tr>
    `).join('');
}

// Filter Users by License
function filterUsersByLicense(license) {
    const filteredUsers = license ? mockData.users.filter(user => user.license === license) : mockData.users;
    
    const tbody = document.getElementById('usersTable');
    tbody.innerHTML = filteredUsers.map(user => `
        <tr>
            <td><strong>${user.name}</strong></td>
            <td>${user.email}</td>
            <td><span class="badge badge-info">${user.license}</span></td>
            <td><span class="badge ${user.mfa ? 'badge-success' : 'badge-warning'}">${user.mfa ? 'Enabled' : 'Disabled'}</span></td>
            <td>${user.lastSignIn}</td>
            <td><span class="badge ${user.status === 'Active' ? 'badge-success' : 'badge-secondary'}">${user.status}</span></td>
            <td>
                <button class="action-btn" onclick="viewUserDetails('${user.email}')">
                    <i class="fas fa-eye"></i> View
                </button>
            </td>
        </tr>
    `).join('');
}

// View User Details (placeholder function)
function viewUserDetails(email) {
    const user = mockData.users.find(u => u.email === email);
    alert(`User Details:\n\nName: ${user.name}\nEmail: ${user.email}\nLicense: ${user.license}\nMFA: ${user.mfa ? 'Enabled' : 'Disabled'}\nDepartment: ${user.department}`);
}

// Report Generation (placeholder functions)
document.querySelectorAll('.report-card .btn').forEach(btn => {
    btn.addEventListener('click', function() {
        const reportType = this.closest('.report-card').querySelector('h3').textContent;
        alert(`Generating ${reportType}...\n\nThis feature would export a comprehensive report in Excel or PDF format.`);
    });
});

// Export functionality
document.querySelectorAll('.btn').forEach(btn => {
    if (btn.textContent.includes('Export')) {
        btn.addEventListener('click', function() {
            alert('Export functionality:\n\nThis would export the current view to Excel, CSV, or PDF format with all the displayed data.');
        });
    }
    
    if (btn.textContent.includes('Refresh')) {
        btn.addEventListener('click', function() {
            alert('Refreshing data...\n\nIn a production environment, this would fetch the latest data from Office 365.');
            location.reload();
        });
    }
});

// ==================== User 360 Profile Modal ====================

// Detailed user profile data
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
        activity: getUserActivityData(user)
    };
}

// Get enabled services based on license
function getEnabledServices(license) {
    const serviceMap = {
        'Office 365 E5': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Yammer', 'Power BI Pro', 'Power Apps', 'Power Automate', 'Stream', 'Forms', 'Planner', 'To-Do', 'Delve', 'Sway', 'Advanced Threat Protection', 'Azure Information Protection'],
        'Office 365 E3': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Yammer', 'Power Apps', 'Power Automate', 'Stream', 'Forms', 'Planner', 'To-Do', 'Delve', 'Sway'],
        'Office 365 E1': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Yammer', 'Forms', 'Planner', 'To-Do'],
        'Business Premium': ['Exchange Online', 'SharePoint Online', 'OneDrive', 'Teams', 'Outlook', 'Word', 'Excel', 'PowerPoint', 'OneNote']
    };
    return serviceMap[license] || [];
}

// Get user groups
function getUserGroups(userName) {
    const userGroupMap = {
        'Sarah Johnson': ['Sales Team', 'Executive Team', 'All Employees'],
        'Michael Chen': ['IT Department', 'All Employees', 'Remote Workers'],
        'Emily Rodriguez': ['Marketing All Hands', 'Project Phoenix', 'All Employees'],
        'David Kim': ['Executive Team', 'All Employees'],
        'Jessica Martinez': ['Executive Team', 'All Employees', 'Remote Workers'],
        'Robert Taylor': ['All Employees'],
        'Amanda White': ['Sales Team', 'All Employees', 'Remote Workers'],
        'James Anderson': ['IT Department', 'All Employees'],
        'Lisa Brown': ['Marketing All Hands', 'All Employees'],
        'Christopher Lee': ['All Employees']
    };
    
    const userGroups = userGroupMap[userName] || [];
    return mockData.groups.filter(g => userGroups.includes(g.name)).map(g => ({
        ...g,
        role: Math.random() > 0.7 ? 'Owner' : 'Member'
    }));
}

// Get mailbox data
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

// Get SharePoint/OneDrive data
function getUserSharePointData(user) {
    return {
        storage: {
            used: (Math.random() * 800 + 100).toFixed(2) + ' GB',
            quota: '1 TB',
            percentage: Math.floor(Math.random() * 60 + 20)
        },
        sites: [
            { name: user.department + ' Team Site', role: 'Member', url: '/sites/' + user.department.toLowerCase() },
            { name: 'Company Intranet', role: 'Visitor', url: '/sites/intranet' },
            { name: 'Projects Hub', role: 'Member', url: '/sites/projects' }
        ],
        sharedFiles: Math.floor(Math.random() * 150 + 50),
        externalSharing: user.license.includes('E5')
    };
}

// Get Teams data
function getUserTeamsData(user) {
    const teamMap = {
        'Sales': ['Sales Team', 'All Company'],
        'IT': ['IT Department', 'All Company', 'Tech Support'],
        'Marketing': ['Marketing Team', 'All Company', 'Creative'],
        'Finance': ['Finance Team', 'All Company'],
        'HR': ['HR Team', 'All Company'],
        'Operations': ['Operations Team', 'All Company'],
        'Support': ['Support Team', 'All Company']
    };
    
    return {
        teams: (teamMap[user.department] || ['All Company']).map(name => ({
            name: name,
            role: Math.random() > 0.8 ? 'Owner' : 'Member',
            channels: Math.floor(Math.random() * 10 + 3)
        })),
        guestAccessEnabled: true,
        externalAccess: user.license.includes('E5')
    };
}

// Get security data
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

// Get activity data
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

// Helper functions
function getRandomPastDate(daysAgo) {
    const date = new Date();
    date.setDate(date.getDate() - Math.floor(Math.random() * daysAgo));
    return date.toISOString().split('T')[0];
}

function generateSignInData() {
    return Array.from({length: 30}, (_, i) => ({
        day: i + 1,
        count: Math.floor(Math.random() * 15 + 5)
    }));
}

// View User Details - Updated Function
function viewUserDetails(email) {
    // Open user profile in new tab
    window.open(`user-profile.html?email=${encodeURIComponent(email)}`, '_blank');
}

// Initialize Profile Tabs
function initializeProfileTabs() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', function() {
            // Remove active class from all buttons
            tabBtns.forEach(b => b.classList.remove('active'));
            this.classList.add('active');

            // Hide all tab contents
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });

            // Show selected tab
            const tabId = this.getAttribute('data-tab') + '-tab';
            document.getElementById(tabId).classList.add('active');
        });
    });
}

// Populate Overview Tab
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

// Populate Licenses Tab
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

// Populate Permissions Tab
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

// Populate Groups Tab
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

// Populate Mailbox Tab
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

// Populate SharePoint Tab
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

// Populate Teams Tab
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

// Populate Security Tab
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

// Populate Activity Tab
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
                    backgroundColor: 'rgba(0, 120, 212, 0.7)',
                    borderColor: '#0078d4',
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
                <div class="usage-fill" style="width: ${percentage}%"></div>
            </div>
            <p>${percentage}%</p>
        </div>
    `).join('');
}

// Close Modal
function closeUser360Modal() {
    const modal = document.getElementById('user360Modal');
    modal.classList.remove('show');
}

// Export User Profile
function exportUserProfile() {
    alert('User Profile Export\n\nThis would generate a comprehensive PDF or Excel report containing:\n\n• User information\n• License details\n• All permissions and roles\n• Group memberships\n• Mailbox statistics\n• Activity logs\n• Security status\n• Compliance information');
}

// Close modal when clicking outside
window.onclick = function(event) {
    const modal = document.getElementById('user360Modal');
    if (event.target === modal) {
        closeUser360Modal();
    }
}

// Initialize Org Chart Filters
function initializeOrgChartFilters() {
    // Department filter for org chart
    const deptFilter = document.getElementById('deptFilter');
    if (deptFilter) {
        deptFilter.addEventListener('change', function() {
            const selectedDept = this.value;
            const orgNodes = document.querySelectorAll('.org-node');
            
            orgNodes.forEach(node => {
                const deptBadge = node.querySelector('.org-node-department');
                if (selectedDept === 'all' || !deptBadge) {
                    node.style.display = 'flex';
                } else {
                    node.style.display = deptBadge.textContent === selectedDept ? 'flex' : 'none';
                }
            });
        });
    }

    // Export org chart button
    const exportOrgChartBtn = document.getElementById('exportOrgChart');
    if (exportOrgChartBtn) {
        exportOrgChartBtn.addEventListener('click', function() {
            alert('Organization chart export functionality will be implemented with a PDF generation library in production.');
        });
    }
}

// Quick Actions Modal Functions
function showClonePermissionsModal(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    const otherUsers = mockData.users.filter(u => u.email !== userEmail);
    
    const userOptions = otherUsers.map(u => `<option value="${u.email}">${u.name}</option>`).join('');
    
    const modalContent = `
        <h3><i class="fas fa-clone"></i> Clone Permissions</h3>
        <p>Clone permissions from <strong>${user.name}</strong> to another user</p>
        <div style="margin: 1.5rem 0;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 500;">Select Target User:</label>
            <select id="cloneTargetUser" style="width: 100%; padding: 0.75rem; border: 1px solid var(--border-color); border-radius: 6px; font-size: 0.95rem;">
                <option value="">-- Choose a user --</option>
                ${userOptions}
            </select>
        </div>
        <div style="margin: 1.5rem 0; padding: 1rem; background: var(--background-alt); border-radius: 6px;">
            <label style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem; cursor: pointer;">
                <input type="checkbox" checked> Clone Group Memberships
            </label>
            <label style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem; cursor: pointer;">
                <input type="checkbox" checked> Clone Admin Roles
            </label>
            <label style="display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.75rem; cursor: pointer;">
                <input type="checkbox" checked> Clone Mailbox Permissions
            </label>
            <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
                <input type="checkbox" checked> Clone Teams Access
            </label>
        </div>
        <div style="display: flex; gap: 1rem; margin-top: 2rem;">
            <button onclick="executeClonePermissions('${userEmail}')" class="btn btn-primary" style="flex: 1;">
                <i class="fas fa-check"></i> Clone Permissions
            </button>
            <button onclick="closeQuickActionModal()" class="btn btn-secondary" style="flex: 1;">
                <i class="fas fa-times"></i> Cancel
            </button>
        </div>
    `;
    
    showQuickActionModal(modalContent);
}

function showQuickActionsModal(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    
    const modalContent = `
        <h3><i class="fas fa-bolt"></i> Quick Actions</h3>
        <p>Quick access actions for <strong>${user.name}</strong></p>
        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 1rem; margin: 1.5rem 0;">
            <button onclick="grantTeamsAccessFromList('${userEmail}')" class="quick-action-card">
                <i class="fas fa-comments" style="font-size: 2rem; color: #464EB8; margin-bottom: 0.5rem;"></i>
                <h4>Teams Access</h4>
                <p>Grant Microsoft Teams access</p>
            </button>
            <button onclick="grantMailboxAccessFromList('${userEmail}')" class="quick-action-card">
                <i class="fas fa-envelope" style="font-size: 2rem; color: #0078D4; margin-bottom: 0.5rem;"></i>
                <h4>Mailbox</h4>
                <p>Grant mailbox permissions</p>
            </button>
            <button onclick="grantSharePointAccessFromList('${userEmail}')" class="quick-action-card">
                <i class="fas fa-folder-open" style="font-size: 2rem; color: #D83B01; margin-bottom: 0.5rem;"></i>
                <h4>SharePoint</h4>
                <p>Grant SharePoint access</p>
            </button>
            <button onclick="addToGroupFromList('${userEmail}')" class="quick-action-card">
                <i class="fas fa-users" style="font-size: 2rem; color: #107C10; margin-bottom: 0.5rem;"></i>
                <h4>Add to Group</h4>
                <p>Add to security group</p>
            </button>
        </div>
        <button onclick="closeQuickActionModal()" class="btn btn-secondary" style="width: 100%; margin-top: 1rem;">
            <i class="fas fa-times"></i> Close
        </button>
    `;
    
    showQuickActionModal(modalContent);
}

function showAssignRoleModal(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    
    const modalContent = `
        <h3><i class="fas fa-user-shield"></i> Assign Admin Role</h3>
        <p>Assign administrative role to <strong>${user.name}</strong></p>
        <div style="display: grid; gap: 1rem; margin: 1.5rem 0;">
            <button onclick="assignRoleFromList('${userEmail}', 'Global Administrator')" class="role-option">
                <i class="fas fa-crown"></i>
                <div>
                    <h4>Global Administrator</h4>
                    <p>Full access to all admin features</p>
                </div>
            </button>
            <button onclick="assignRoleFromList('${userEmail}', 'Exchange Administrator')" class="role-option">
                <i class="fas fa-envelope"></i>
                <div>
                    <h4>Exchange Administrator</h4>
                    <p>Manage Exchange Online settings</p>
                </div>
            </button>
            <button onclick="assignRoleFromList('${userEmail}', 'SharePoint Administrator')" class="role-option">
                <i class="fas fa-sitemap"></i>
                <div>
                    <h4>SharePoint Administrator</h4>
                    <p>Manage SharePoint sites and settings</p>
                </div>
            </button>
            <button onclick="assignRoleFromList('${userEmail}', 'Teams Administrator')" class="role-option">
                <i class="fas fa-comments"></i>
                <div>
                    <h4>Teams Administrator</h4>
                    <p>Manage Microsoft Teams settings</p>
                </div>
            </button>
            <button onclick="assignRoleFromList('${userEmail}', 'User Administrator')" class="role-option">
                <i class="fas fa-users-cog"></i>
                <div>
                    <h4>User Administrator</h4>
                    <p>Manage users and groups</p>
                </div>
            </button>
        </div>
        <button onclick="closeQuickActionModal()" class="btn btn-secondary" style="width: 100%; margin-top: 1rem;">
            <i class="fas fa-times"></i> Cancel
        </button>
    `;
    
    showQuickActionModal(modalContent);
}

function showQuickActionModal(content) {
    // Remove existing modal if any
    const existingModal = document.getElementById('quickActionModal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Create modal
    const modal = document.createElement('div');
    modal.id = 'quickActionModal';
    modal.className = 'modal-overlay';
    modal.innerHTML = `
        <div class="modal-content-quick">
            ${content}
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Close on outside click
    modal.onclick = function(event) {
        if (event.target === modal) {
            closeQuickActionModal();
        }
    };
}

function closeQuickActionModal() {
    const modal = document.getElementById('quickActionModal');
    if (modal) {
        modal.remove();
    }
}

function executeClonePermissions(sourceEmail) {
    const targetEmail = document.getElementById('cloneTargetUser').value;
    if (!targetEmail) {
        alert('Please select a target user.');
        return;
    }
    
    const sourceUser = mockData.users.find(u => u.email === sourceEmail);
    const targetUser = mockData.users.find(u => u.email === targetEmail);
    
    alert(`✅ Permissions Cloned Successfully!\n\nFrom: ${sourceUser.name}\nTo: ${targetUser.name}\n\nThe following have been replicated:\n✓ Group memberships\n✓ Admin roles\n✓ Mailbox permissions\n✓ Teams access`);
    closeQuickActionModal();
}

function grantTeamsAccessFromList(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    alert(`✅ Teams Access Granted\n\nUser: ${user.name}\n\n✓ Teams license enabled\n✓ Added to default teams\n✓ Chat and calling configured`);
    closeQuickActionModal();
}

function grantMailboxAccessFromList(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    alert(`📧 Grant Mailbox Permissions\n\nFor: ${user.name}\n\nSelect:\n• Mailbox to grant access to\n• Permission type (Full Access, Send As, etc.)`);
}

function grantSharePointAccessFromList(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    alert(`📁 Grant SharePoint Access\n\nFor: ${user.name}\n\nSelect:\n• SharePoint site\n• Permission level (Read, Edit, Full Control)`);
}

function addToGroupFromList(userEmail) {
    const user = mockData.users.find(u => u.email === userEmail);
    alert(`👥 Add to Security Group\n\nFor: ${user.name}\n\nAvailable groups:\n• Executive Team\n• Sales Team\n• Production Team\n• Operations Team\n• Technical Team`);
}

function assignRoleFromList(userEmail, roleName) {
    const user = mockData.users.find(u => u.email === userEmail);
    alert(`✅ Admin Role Assigned\n\nRole: ${roleName}\nUser: ${user.name}\n\n✓ Administrative privileges granted\n✓ Access to admin portal enabled`);
    closeQuickActionModal();
}
