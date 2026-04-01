/**
 * Distributor Product Lookup Widget
 * For Zoho CRM Quotes module integration
 * Version: 3.5
 * Updated: March 31, 2026 — Fix TD Synnex manufacturer RPC field mapping (manufacturer_name)
 * Supports: TD Synnex, Ingram Micro, ADI Global
 * Features: Single & Bulk search modes, MSRP comparison, manufacturer resolution,
 *           customer discount %, smart column auto-mapping, lazy API manufacturer verification,
 *           scroll-to-focus panel transitions, product details loading UX, admin panel
 */

// =====================================================
// CONFIGURATION
// =====================================================
const PROXY_BASE = 'https://yasocskmsepalujntkau.supabase.co/functions/v1/ingram-proxy';
const TDSYNNEX_BASE = 'https://yasocskmsepalujntkau.supabase.co/functions/v1';
const TDSYNNEX_PROXY_BASE = 'https://yasocskmsepalujntkau.supabase.co/functions/v1/tdsynnex-proxy';
const ADI_PROXY_BASE = 'https://yasocskmsepalujntkau.supabase.co/functions/v1/adi-proxy';
const SUPABASE_URL = 'https://yasocskmsepalujntkau.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inlhc29jc2ttc2VwYWx1am50a2F1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjkyMTIyNzUsImV4cCI6MjA4NDc4ODI3NX0.81g9I9iw1EbNOVCgOtNDTZtDavFMUbiprUMriSFd34c';
const PAGE_SIZE = 50;

// Distributor configurations
const DISTRIBUTORS = {
    ingram: {
        name: 'Ingram Micro',
        apiPrefix: '/api',
        color: '#0ea5e9'
    },
    tdsynnex: {
        name: 'TD Synnex',
        apiPrefix: '/tdsynnex',
        color: '#10b981'
    },
    adi: {
        name: 'ADI Global',
        apiPrefix: '/adi',
        color: '#8b5cf6'
    }
};

// =====================================================
// STATE MANAGEMENT
// =====================================================
const state = {
    currentDistributor: 'ingram',
    // Filters
    manufacturer: '',
    category: '',
    subcategory: '',
    cat3: '',  // TD Synnex category level 3
    skuType: '',
    skuKeyword: '',
    // Filter loading state
    loadingFilters: {
        category: false,
        subcategory: false,
        cat3: false,
        skuType: false
    },
    filterParams: {
        category: '',
        subcategory: '',
        cat3: '',
        skuType: ''
    },
    // Pagination and products
    currentPage: 1,
    totalRecords: 0,
    totalPages: 1,
    selectedProducts: new Map(),
    queuedProducts: [],
    groupByManufacturer: true,
    pricingMode: 'msrp',
    isAuthenticated: false,
    pendingResponseId: null,
    parentContext: null,
    currentProducts: [],
    pricingData: {},
    rawApiVisible: false,
    // SKU-first search mode
    skuSearchMode: false,
    pendingSkuFilter: '',
    skuManufacturerOptions: [],
    // Manufacturer Resolution Panel State (Phase 3)
    prefetchedManufacturers: [],
    mfrResolutions: new Map(),
    unresolvedManufacturers: [],
    mfrResolutionPromise: null,
    manufacturerMappingsData: [],  // Cached mappings from Supabase
    searchMode: 'single',      // 'single' or 'bulk'
    verifiedIngramMfrs: new Set(),  // Session-level tracking for lazy API verification
    // Admin
    adminClickTimes: [],
    adminPanelOpen: false,
    currentAdminPage: 'mfr-filters',
    // Admin - Manufacturer Filters
    adminFilterData: {},
    adminActiveTab: 'tdsynnex',
    adminPending: {},
    adminFilterLoading: false,
    adminExpandedGroups: new Set(),
    adminLetterFilter: 'All',
    adminLetterScope: 'available',
    // Admin - Workflow Import
    workflowState: {
        tdsynnex: { running: false, runId: null, status: null, conclusion: null },
        ingram: { running: false, runId: null, status: null, conclusion: null },
        adi: { running: false, runId: null, status: null, conclusion: null },
    },
    workflowPollingTimers: {},
    workflowLastRun: {},
};

let searchTimeout = null;
let draggedItem = null;
let draggedGroup = null;
let resizeStartY = 0;
let resizeStartHeight = 0;
let isResizing = false;
let isResizingQueue = false;
let queueResizeStartX = 0;
let queueResizeStartWidth = 0;

// =====================================================
// ADMIN PANEL
// =====================================================

/**
 * Triple-click handler for admin cog — requires 3 clicks within 600ms
 */
function handleAdminCogClick() {
    const now = Date.now();
    state.adminClickTimes.push(now);

    // Keep only clicks within last 600ms
    state.adminClickTimes = state.adminClickTimes.filter(t => now - t < 600);

    if (state.adminClickTimes.length >= 3) {
        state.adminClickTimes = [];
        openAdminPanel();
    }
}

function openAdminPanel() {
    document.getElementById('adminPanel').style.display = 'flex';
    document.querySelector('.content-wrapper').style.display = 'none';
    document.querySelector('.header-buttons').style.display = 'none';
    document.getElementById('adminCogBtn').style.display = 'none';
    state.adminPanelOpen = true;
    // Auto-load filter data for current admin tab
    if (state.currentAdminPage === 'mfr-filters') {
        loadMfrFilterData(state.adminActiveTab);
    }
}

function closeAdminPanel() {
    document.getElementById('adminPanel').style.display = 'none';
    document.querySelector('.content-wrapper').style.display = '';
    document.querySelector('.header-buttons').style.display = '';
    document.getElementById('adminCogBtn').style.display = '';
    state.adminPanelOpen = false;
}

function selectAdminPage(pageId) {
    // Update nav
    document.querySelectorAll('.admin-nav-item').forEach(item => {
        item.classList.toggle('active', item.dataset.adminPage === pageId);
    });

    // Update content
    document.querySelectorAll('.admin-page').forEach(page => {
        page.classList.toggle('active', page.id === `adminPage-${pageId}`);
    });

    state.currentAdminPage = pageId;

    // Auto-load data when navigating to mfr-filters page
    if (pageId === 'mfr-filters') {
        loadMfrFilterData(state.adminActiveTab);
    }
}

// =====================================================
// ADMIN — MANUFACTURER FILTER MANAGEMENT
// =====================================================

const GITHUB_PROXY_BASE = `${SUPABASE_URL}/functions/v1/github-proxy`;

/**
 * Select a distributor tab in the admin filter page
 */
function selectMfrAdminTab(dist) {
    state.adminActiveTab = dist;
    state.adminExpandedGroups.clear();
    state.adminLetterFilter = 'All';
    // Update tab UI
    document.querySelectorAll('.mfr-admin-tab').forEach(t => {
        t.classList.toggle('active', t.dataset.dist === dist);
    });
    // Clear search inputs
    const availSearch = document.getElementById('mfrAvailSearch');
    const inclSearch = document.getElementById('mfrIncludedSearch');
    if (availSearch) availSearch.value = '';
    if (inclSearch) inclSearch.value = '';
    // Load fresh data (no caching)
    loadMfrFilterData(dist);
}

/**
 * Load manufacturer filter data from the github-proxy edge function
 */
async function loadMfrFilterData(dist) {
    state.adminFilterLoading = true;
    showMfrLoading(true);
    showMfrColumns(false);
    showMfrEmpty(false);

    try {
        const res = await fetch(`${GITHUB_PROXY_BASE}?action=get-filters&distributor=${dist}`);
        if (res.status === 404 || !res.ok) {
            // No filter data yet — show empty state
            state.adminFilterData[dist] = null;
            showMfrLoading(false);
            showMfrEmpty(true);
            renderMfrStats(dist, null);
            updateMfrPendingBar();
            return;
        }
        const data = await res.json();
        if (data.error) {
            state.adminFilterData[dist] = null;
            showMfrLoading(false);
            showMfrEmpty(true);
            renderMfrStats(dist, null);
            updateMfrPendingBar();
            return;
        }
        state.adminFilterData[dist] = data;

        // Initialize pending if not already
        if (!state.adminPending[dist]) {
            state.adminPending[dist] = { additions: new Set(), removals: new Set() };
        }

        showMfrLoading(false);
        showMfrColumns(true);
        renderMfrStats(dist, data);
        renderLetterBar();
        renderMfrColumns();
        updateMfrPendingBar();
    } catch (err) {
        console.error('Failed to load mfr filter data:', err);
        state.adminFilterData[dist] = null;
        showMfrLoading(false);
        showMfrEmpty(true);
        renderMfrStats(dist, null);
    } finally {
        state.adminFilterLoading = false;
    }
}

function showMfrLoading(show) {
    document.getElementById('mfrAdminLoading').style.display = show ? 'flex' : 'none';
}
function showMfrColumns(show) {
    document.getElementById('mfrAdminColumns').style.display = show ? 'grid' : 'none';
}
function showMfrEmpty(show) {
    document.getElementById('mfrAdminEmpty').style.display = show ? 'flex' : 'none';
}

/**
 * Render stats row for a distributor
 */
function renderMfrStats(dist, data) {
    const el = document.getElementById('mfrAdminStats');
    if (!data) {
        el.innerHTML = '';
        return;
    }
    const s = data.stats || {};
    const pending = state.adminPending[dist] || { additions: new Set(), removals: new Set() };
    const activeCount = (data.active_manufacturers || []).length + pending.additions.size - pending.removals.size;

    // Compute filtered SKUs dynamically
    const activeSet = getEffectiveActiveSet(dist);
    const details = data.manufacturer_details || {};
    let filteredSkus = 0;
    for (const name of activeSet) {
        filteredSkus += (details[name]?.sku_count || 0);
    }

    // For distributors without total_skus in stats, compute from manufacturer_details
    let totalSkus = s.total_skus || s.total_skus_in_file || 0;
    if (!totalSkus && details) {
        for (const name of Object.keys(details)) {
            totalSkus += (details[name]?.sku_count || 0);
        }
    }

    el.innerHTML = `
        <div class="mfr-admin-stat">
            <span class="mfr-admin-stat-label">Known</span>
            <span class="mfr-admin-stat-val">${fmtNum(s.total_manufacturers || s.total_known || (data.all_known_manufacturers || []).length)}</span>
        </div>
        <div class="mfr-admin-stat-separator"></div>
        <div class="mfr-admin-stat">
            <span class="mfr-admin-stat-label">Active</span>
            <span class="mfr-admin-stat-val mfr-admin-stat-val--accent">${fmtNum(Math.max(0, activeCount))}</span>
        </div>
        <div class="mfr-admin-stat-separator"></div>
        <div class="mfr-admin-stat">
            <span class="mfr-admin-stat-label">Total SKUs</span>
            <span class="mfr-admin-stat-val">${fmtNum(totalSkus)}</span>
        </div>
        <div class="mfr-admin-stat-separator"></div>
        <div class="mfr-admin-stat">
            <span class="mfr-admin-stat-label">Filtered SKUs</span>
            <span class="mfr-admin-stat-val">${fmtNum(filteredSkus)}</span>
        </div>
        <div class="mfr-admin-stat-separator"></div>
        <div class="mfr-admin-stat">
            <span class="mfr-admin-stat-label">Last Run</span>
            <span class="mfr-admin-stat-val">${s.last_run ? fmtRelDate(s.last_run) : '--'}</span>
        </div>
    `;
}

/**
 * Get the effective active set (original + pending additions - pending removals)
 */
function getEffectiveActiveSet(dist) {
    const data = state.adminFilterData[dist];
    if (!data) return new Set();
    const active = new Set(data.active_manufacturers || []);
    const pending = state.adminPending[dist] || { additions: new Set(), removals: new Set() };
    for (const n of pending.additions) active.add(n);
    for (const n of pending.removals) active.delete(n);
    return active;
}

/**
 * Format number with commas
 */
function fmtNum(n) {
    return (n || 0).toLocaleString();
}

/**
 * Format a date string as relative or absolute
 */
function fmtRelDate(dateStr) {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return '--';
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) + ' ' +
           d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });
}

/**
 * LCP grouping for Ingram/ADI manufacturer names
 * Groups manufacturers that share a 2+ word prefix
 */
function groupManufacturers(names, details) {
    if (names.length === 0) return [];

    const sorted = [...names].sort();
    const groups = [];
    let i = 0;

    while (i < sorted.length) {
        // Look ahead to find names sharing a common prefix
        let bestPrefix = null;
        let groupEnd = i;

        for (let j = i + 1; j < sorted.length; j++) {
            const prefix = sharedWordPrefix(sorted[i], sorted[j]);
            if (prefix && prefix.split(/\s+/).length >= 2) {
                // Check if all names from i to j share this prefix
                let allShare = true;
                for (let k = i; k <= j; k++) {
                    if (!sorted[k].toUpperCase().startsWith(prefix.toUpperCase())) {
                        allShare = false;
                        break;
                    }
                }
                if (allShare) {
                    bestPrefix = prefix;
                    groupEnd = j;
                } else {
                    break;
                }
            } else {
                break;
            }
        }

        if (bestPrefix && groupEnd > i) {
            // Multi-name group
            const members = sorted.slice(i, groupEnd + 1);
            let totalSkus = 0;
            for (const m of members) totalSkus += (details[m]?.sku_count || 0);
            groups.push({
                type: 'group',
                prefix: bestPrefix,
                members: members,
                totalSkus: totalSkus,
            });
            i = groupEnd + 1;
        } else {
            // Single item
            groups.push({
                type: 'single',
                name: sorted[i],
                skuCount: details[sorted[i]]?.sku_count || 0,
            });
            i++;
        }
    }

    return groups;
}

/**
 * Find shared word-level prefix between two strings
 */
function sharedWordPrefix(a, b) {
    const wordsA = a.split(/\s+/);
    const wordsB = b.split(/\s+/);
    const common = [];
    for (let i = 0; i < Math.min(wordsA.length, wordsB.length); i++) {
        if (wordsA[i].toUpperCase() === wordsB[i].toUpperCase()) {
            common.push(wordsA[i]);
        } else {
            break;
        }
    }
    return common.length >= 2 ? common.join(' ') : null;
}

/**
 * Check if a distributor uses grouped display (Ingram/ADI have unclean names)
 */
function isGroupedDistributor(dist) {
    return dist === 'ingram' || dist === 'adi';
}

/**
 * Render both Available and Included columns
 */
function renderMfrColumns() {
    const dist = state.adminActiveTab;
    const data = state.adminFilterData[dist];
    if (!data) return;

    const activeSet = getEffectiveActiveSet(dist);
    const pending = state.adminPending[dist] || { additions: new Set(), removals: new Set() };
    const details = data.manufacturer_details || {};
    const allNames = data.all_known_manufacturers || [];

    const availSearch = (document.getElementById('mfrAvailSearch')?.value || '').toUpperCase();
    const inclSearch = (document.getElementById('mfrIncludedSearch')?.value || '').toUpperCase();

    // Split into available and included
    const availNames = allNames.filter(n => !activeSet.has(n));
    const inclNames = allNames.filter(n => activeSet.has(n));

    // Apply letter filter
    const letterFilter = state.adminLetterFilter;
    const scope = state.adminLetterScope;
    const applyLetter = letterFilter && letterFilter !== 'All';
    const letterAvail = (applyLetter && (scope === 'available' || scope === 'both')) ? availNames.filter(n => n.charAt(0).toUpperCase() === letterFilter) : availNames;
    const letterIncl = (applyLetter && (scope === 'included' || scope === 'both')) ? inclNames.filter(n => n.charAt(0).toUpperCase() === letterFilter) : inclNames;

    // Apply search filter
    const filteredAvail = availSearch ? letterAvail.filter(n => n.toUpperCase().includes(availSearch)) : letterAvail;
    const filteredIncl = inclSearch ? letterIncl.filter(n => n.toUpperCase().includes(inclSearch)) : letterIncl;

    // Render columns
    const availList = document.getElementById('mfrAvailList');
    const inclList = document.getElementById('mfrIncludedList');

    if (isGroupedDistributor(dist)) {
        availList.innerHTML = renderGroupedList(filteredAvail, details, 'avail', pending);
        inclList.innerHTML = renderGroupedList(filteredIncl, details, 'included', pending);
    } else {
        availList.innerHTML = renderFlatList(filteredAvail, details, 'avail', pending);
        inclList.innerHTML = renderFlatList(filteredIncl, details, 'included', pending);
    }

    // Update counts
    document.getElementById('mfrAvailCount').textContent = filteredAvail.length;
    document.getElementById('mfrIncludedCount').textContent = filteredIncl.length;

    // Update stats (filtered SKUs change with pending changes)
    renderMfrStats(dist, data);
    updateMfrPendingBar();
}

/**
 * Render a flat list (for TD Synnex)
 */
function renderFlatList(names, details, side, pending) {
    if (names.length === 0) {
        return '<div class="mfr-admin-list-empty">No manufacturers found</div>';
    }
    const sorted = [...names].sort();
    const arrowSvg = side === 'avail'
        ? '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"></polyline></svg>'
        : '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"></polyline></svg>';

    // Separate pending items from non-pending items
    const pendingItems = [];
    const normalItems = [];
    sorted.forEach(name => {
        const isPendingAdd = pending.additions.has(name);
        const isPendingRemove = pending.removals.has(name);
        if ((side === 'included' && isPendingAdd) || (side === 'avail' && isPendingRemove)) {
            pendingItems.push(name);
        } else {
            normalItems.push(name);
        }
    });

    const renderItem = (name, showBadge) => {
        const sku = details[name]?.sku_count || 0;
        const pendClass = pending.additions.has(name) ? ' mfr-admin-item--pending-add'
                        : pending.removals.has(name) ? ' mfr-admin-item--pending-remove' : '';
        const action = side === 'avail' ? `mfrIncludeName('${escAttr(name)}')` : `mfrExcludeName('${escAttr(name)}')`;
        const badge = showBadge
            ? (side === 'included'
                ? '<span class="mfr-admin-pending-badge mfr-admin-pending-badge--add">+</span>'
                : '<span class="mfr-admin-pending-badge mfr-admin-pending-badge--remove">\u2212</span>')
            : '';
        return `<div class="mfr-admin-item${pendClass}" onclick="${action}">
            ${badge}<span class="mfr-admin-item-name" title="${escAttr(name)}">${escHtml(name)}</span>
            <span class="mfr-admin-item-sku">${fmtNum(sku)}</span>
            <span class="mfr-admin-item-arrow">${arrowSvg}</span>
        </div>`;
    };

    let html = '';
    if (pendingItems.length > 0) {
        html += pendingItems.map(n => renderItem(n, true)).join('');
        if (normalItems.length > 0) {
            html += '<div class="mfr-admin-pending-separator">All Manufacturers</div>';
        }
    }
    html += normalItems.map(n => renderItem(n, false)).join('');
    return html;
}

/**
 * Render a grouped list (for Ingram/ADI)
 */
function renderGroupedList(names, details, side, pending) {
    if (names.length === 0) {
        return '<div class="mfr-admin-list-empty">No manufacturers found</div>';
    }
    const groups = groupManufacturers(names, details);
    const arrowSvg = side === 'avail'
        ? '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"></polyline></svg>'
        : '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"></polyline></svg>';
    const chevronSvg = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"></polyline></svg>';

    const renderSingle = (g, showBadge) => {
        const pendClass = pending.additions.has(g.name) ? ' mfr-admin-item--pending-add'
                        : pending.removals.has(g.name) ? ' mfr-admin-item--pending-remove' : '';
        const action = side === 'avail' ? `mfrIncludeName('${escAttr(g.name)}')` : `mfrExcludeName('${escAttr(g.name)}')`;
        const badge = showBadge
            ? (side === 'included'
                ? '<span class="mfr-admin-pending-badge mfr-admin-pending-badge--add">+</span>'
                : '<span class="mfr-admin-pending-badge mfr-admin-pending-badge--remove">\u2212</span>')
            : '';
        return `<div class="mfr-admin-item${pendClass}" onclick="${action}">
            ${badge}<span class="mfr-admin-item-name" title="${escAttr(g.name)}">${escHtml(g.name)}</span>
            <span class="mfr-admin-item-sku">${fmtNum(g.skuCount)}</span>
            <span class="mfr-admin-item-arrow">${arrowSvg}</span>
        </div>`;
    };

    const renderGroup = (g, showBadge) => {
        const groupId = `grp-${side}-${escAttr(g.prefix).replace(/\s+/g, '-')}`;
        const isExpanded = state.adminExpandedGroups.has(groupId);

        // Determine group pending state
        const allPendingAdd = g.members.every(m => pending.additions.has(m));
        const allPendingRemove = g.members.every(m => pending.removals.has(m));
        const anyPending = g.members.some(m => pending.additions.has(m) || pending.removals.has(m));
        let groupPendClass = '';
        if (allPendingAdd) groupPendClass = ' mfr-admin-group--pending-add';
        else if (allPendingRemove) groupPendClass = ' mfr-admin-group--pending-remove';
        else if (anyPending) groupPendClass = ' mfr-admin-group--partial';

        const badge = showBadge
            ? (side === 'included'
                ? '<span class="mfr-admin-pending-badge mfr-admin-pending-badge--add">+</span>'
                : '<span class="mfr-admin-pending-badge mfr-admin-pending-badge--remove">\u2212</span>')
            : '';

        const groupAction = side === 'avail' ? `mfrIncludeGroup(event, ${JSON.stringify(g.members).replace(/"/g, '&quot;')})` : `mfrExcludeGroup(event, ${JSON.stringify(g.members).replace(/"/g, '&quot;')})`;

        const children = g.members.map(name => {
            const sku = details[name]?.sku_count || 0;
            const pendClass = pending.additions.has(name) ? ' mfr-admin-item--pending-add'
                            : pending.removals.has(name) ? ' mfr-admin-item--pending-remove' : '';
            const action = side === 'avail' ? `mfrIncludeName('${escAttr(name)}')` : `mfrExcludeName('${escAttr(name)}')`;
            const suffix = name.substring(g.prefix.length).trim() || name;
            return `<div class="mfr-admin-item${pendClass}" onclick="event.stopPropagation(); ${action}">
                <span class="mfr-admin-item-name" title="${escAttr(name)}">${escHtml(suffix)}</span>
                <span class="mfr-admin-item-sku">${fmtNum(sku)}</span>
                <span class="mfr-admin-item-arrow">${arrowSvg}</span>
            </div>`;
        }).join('');

        return `<div class="mfr-admin-group${groupPendClass}">
            <div class="mfr-admin-group-header" onclick="${groupAction}">
                ${badge}<span class="mfr-admin-group-chevron${isExpanded ? ' expanded' : ''}" onclick="event.stopPropagation(); toggleMfrGroup('${groupId}')">
                    ${chevronSvg}
                </span>
                <span class="mfr-admin-group-name" title="${escAttr(g.prefix)}">${escHtml(g.prefix)}</span>
                <span class="mfr-admin-group-meta">
                    <span class="mfr-admin-group-badge">${g.members.length} variants</span>
                    <span class="mfr-admin-group-sku">${fmtNum(g.totalSkus)}</span>
                </span>
                <span class="mfr-admin-group-arrow">${arrowSvg}</span>
            </div>
            <div class="mfr-admin-group-children${isExpanded ? ' open' : ''}" id="${groupId}">
                ${children}
            </div>
        </div>`;
    };

    // Separate groups into pending and non-pending
    const pendingGroups = [];
    const normalGroups = [];
    groups.forEach(g => {
        if (g.type === 'single') {
            const isPendingAdd = pending.additions.has(g.name);
            const isPendingRemove = pending.removals.has(g.name);
            if ((side === 'included' && isPendingAdd) || (side === 'avail' && isPendingRemove)) {
                pendingGroups.push(g);
            } else {
                normalGroups.push(g);
            }
        } else {
            // For groups: only float if ALL members are pending in the relevant direction
            const allPendingAdd = g.members.every(m => pending.additions.has(m));
            const allPendingRemove = g.members.every(m => pending.removals.has(m));
            if ((side === 'included' && allPendingAdd) || (side === 'avail' && allPendingRemove)) {
                pendingGroups.push(g);
            } else {
                normalGroups.push(g);
            }
        }
    });

    let html = '';
    if (pendingGroups.length > 0) {
        html += pendingGroups.map(g => g.type === 'single' ? renderSingle(g, true) : renderGroup(g, true)).join('');
        if (normalGroups.length > 0) {
            html += '<div class="mfr-admin-pending-separator">All Manufacturers</div>';
        }
    }
    html += normalGroups.map(g => g.type === 'single' ? renderSingle(g, false) : renderGroup(g, false)).join('');
    return html;
}

/**
 * Toggle group expansion
 */
function toggleMfrGroup(groupId) {
    const el = document.getElementById(groupId);
    if (!el) return;
    const chevron = el.previousElementSibling?.querySelector('.mfr-admin-group-chevron');
    if (el.classList.contains('open')) {
        el.classList.remove('open');
        if (chevron) chevron.classList.remove('expanded');
        state.adminExpandedGroups.delete(groupId);
    } else {
        el.classList.add('open');
        if (chevron) chevron.classList.add('expanded');
        state.adminExpandedGroups.add(groupId);
    }
}

/**
 * Include a single manufacturer name
 */
function mfrIncludeName(name) {
    const dist = state.adminActiveTab;
    ensurePending(dist);
    const p = state.adminPending[dist];
    const data = state.adminFilterData[dist];
    const origActive = new Set(data?.active_manufacturers || []);

    if (origActive.has(name)) {
        // Was originally active — remove from removals
        p.removals.delete(name);
    } else {
        // New addition
        p.additions.add(name);
    }
    renderMfrColumns();
}

/**
 * Exclude a single manufacturer name
 */
function mfrExcludeName(name) {
    const dist = state.adminActiveTab;
    ensurePending(dist);
    const p = state.adminPending[dist];
    const data = state.adminFilterData[dist];
    const origActive = new Set(data?.active_manufacturers || []);

    if (origActive.has(name)) {
        // Was originally active — add to removals
        p.removals.add(name);
    } else {
        // Was a pending addition — remove it
        p.additions.delete(name);
    }
    renderMfrColumns();
}

/**
 * Include all members of a group
 */
function mfrIncludeGroup(event, members) {
    event.stopPropagation();
    const dist = state.adminActiveTab;
    ensurePending(dist);
    const p = state.adminPending[dist];
    const data = state.adminFilterData[dist];
    const origActive = new Set(data?.active_manufacturers || []);

    for (const name of members) {
        if (origActive.has(name)) {
            p.removals.delete(name);
        } else {
            p.additions.add(name);
        }
    }
    renderMfrColumns();
}

/**
 * Exclude all members of a group
 */
function mfrExcludeGroup(event, members) {
    event.stopPropagation();
    const dist = state.adminActiveTab;
    ensurePending(dist);
    const p = state.adminPending[dist];
    const data = state.adminFilterData[dist];
    const origActive = new Set(data?.active_manufacturers || []);

    for (const name of members) {
        if (origActive.has(name)) {
            p.removals.add(name);
        } else {
            p.additions.delete(name);
        }
    }
    renderMfrColumns();
}

/**
 * Add all visible available manufacturers
 */
function mfrAddAllVisible() {
    const dist = state.adminActiveTab;
    const data = state.adminFilterData[dist];
    if (!data) return;

    ensurePending(dist);
    const activeSet = getEffectiveActiveSet(dist);
    const allNames = data.all_known_manufacturers || [];
    const availNames = allNames.filter(n => !activeSet.has(n));
    const search = (document.getElementById('mfrAvailSearch')?.value || '').toUpperCase();
    const filtered = search ? availNames.filter(n => n.toUpperCase().includes(search)) : availNames;

    const p = state.adminPending[dist];
    const origActive = new Set(data.active_manufacturers || []);

    for (const name of filtered) {
        if (origActive.has(name)) {
            p.removals.delete(name);
        } else {
            p.additions.add(name);
        }
    }
    renderMfrColumns();
}

/**
 * Remove all visible included manufacturers
 */
function mfrRemoveAllVisible() {
    const dist = state.adminActiveTab;
    const data = state.adminFilterData[dist];
    if (!data) return;

    ensurePending(dist);
    const activeSet = getEffectiveActiveSet(dist);
    const allNames = data.all_known_manufacturers || [];
    const inclNames = allNames.filter(n => activeSet.has(n));
    const search = (document.getElementById('mfrIncludedSearch')?.value || '').toUpperCase();
    const filtered = search ? inclNames.filter(n => n.toUpperCase().includes(search)) : inclNames;

    const p = state.adminPending[dist];
    const origActive = new Set(data.active_manufacturers || []);

    for (const name of filtered) {
        if (origActive.has(name)) {
            p.removals.add(name);
        } else {
            p.additions.delete(name);
        }
    }
    renderMfrColumns();
}

/**
 * Ensure pending state exists for distributor
 */
function ensurePending(dist) {
    if (!state.adminPending[dist]) {
        state.adminPending[dist] = { additions: new Set(), removals: new Set() };
    }
}

/**
 * Update the pending changes bar visibility and summary
 */
function updateMfrPendingBar() {
    const dist = state.adminActiveTab;
    const p = state.adminPending[dist];
    const bar = document.getElementById('adminHeaderActions');
    const summary = document.getElementById('adminHeaderPendingText');
    const dot = bar ? bar.querySelector('.admin-header-pending-dot') : null;
    const discardBtn = bar ? bar.querySelector('.admin-header-btn-discard') : null;
    const saveBtn = bar ? bar.querySelector('.admin-header-btn-save') : null;

    if (!bar) return;

    bar.style.display = 'flex';

    const hasChanges = p && (p.additions.size > 0 || p.removals.size > 0);

    if (!hasChanges) {
        if (discardBtn) discardBtn.disabled = true;
        if (saveBtn) saveBtn.disabled = true;
        if (dot) dot.style.display = 'none';
        summary.textContent = 'No pending changes';
        return;
    }

    if (discardBtn) discardBtn.disabled = false;
    if (saveBtn) saveBtn.disabled = false;
    if (dot) dot.style.display = '';
    const parts = [];
    if (p.additions.size > 0) parts.push(`${p.additions.size} addition${p.additions.size !== 1 ? 's' : ''}`);
    if (p.removals.size > 0) parts.push(`${p.removals.size} removal${p.removals.size !== 1 ? 's' : ''}`);
    summary.textContent = parts.join(', ');
}

/**
 * Discard all pending changes for the current tab
 */
function mfrDiscardChanges() {
    const dist = state.adminActiveTab;
    state.adminPending[dist] = { additions: new Set(), removals: new Set() };
    renderMfrColumns();
}

/**
 * Save pending changes via the github-proxy edge function
 */
async function mfrSaveChanges() {
    const dist = state.adminActiveTab;
    const p = state.adminPending[dist];
    if (!p || (p.additions.size === 0 && p.removals.size === 0)) return;

    const saveBtn = document.querySelector('.admin-header-btn-save');
    const origText = saveBtn.innerHTML;
    saveBtn.innerHTML = '<div class="mfr-admin-spinner" style="width:14px;height:14px;border-width:2px;"></div> Saving...';
    saveBtn.disabled = true;

    try {
        const res = await fetch(`${GITHUB_PROXY_BASE}?action=save-filters`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                distributor: dist,
                additions: [...p.additions],
                removals: [...p.removals],
            }),
        });
        const result = await res.json();
        if (result.error) {
            console.error('Save failed:', result.error);
            saveBtn.innerHTML = origText;
            saveBtn.disabled = false;
            return;
        }

        // Optimistic update: apply pending changes to local state
        const effectiveActive = [...getEffectiveActiveSet(dist)];
        state.adminFilterData[dist].active_manufacturers = effectiveActive;

        // Clear pending
        state.adminPending[dist] = { additions: new Set(), removals: new Set() };
        updateMfrPendingBar();
        renderMfrColumns();
        renderMfrStats(dist, state.adminFilterData[dist]);

        // Show success modal
        document.getElementById('mfrSaveModal').style.display = 'flex';

        saveBtn.innerHTML = origText;
        saveBtn.disabled = false;
    } catch (err) {
        console.error('Save error:', err);
        saveBtn.innerHTML = origText;
        saveBtn.disabled = false;
    }
}

/**
 * Cancel — dismiss the save success modal without re-rendering
 */
function mfrCancelModal() {
    document.getElementById('mfrSaveModal').style.display = 'none';
}

/**
 * Dismiss the save success modal
 */
function mfrDismissModal() {
    document.getElementById('mfrSaveModal').style.display = 'none';
    renderMfrColumns();
}

async function mfrApplyNow() {
    document.getElementById('mfrSaveModal').style.display = 'none';
    const dist = state.adminActiveTab;
    const wf = state.workflowState[dist];
    if (wf.running) return;

    wf.running = true;
    wf.status = 'dispatching';
    wf.conclusion = null;
    wf.runId = null;

    try {
        const res = await fetch(`${GITHUB_PROXY_BASE}?action=dispatch-workflow`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ distributor: dist }),
        });
        const result = await res.json();

        if (!result.success) {
            wf.running = false;
            wf.status = 'error';
            showWorkflowToast(dist, false, result.error || 'Dispatch failed');
            return;
        }

        if (result.run_id) {
            wf.runId = result.run_id;
            wf.status = result.status || 'queued';
        } else {
            await new Promise(r => setTimeout(r, 3000));
            const foundId = await findLatestRunId(dist);
            if (foundId) {
                wf.runId = foundId;
                wf.status = 'queued';
            } else {
                wf.status = 'queued';
            }
        }

        startWorkflowPolling(dist);
        showWorkflowRunningToast(dist);
    } catch (err) {
        console.error('Workflow dispatch error:', err);
        wf.running = false;
        wf.status = 'error';
        showWorkflowToast(dist, false, 'Network error dispatching workflow');
    }
}

// ── Workflow Dispatch & Monitoring (Phase 9.3) ──────────────────────────

async function findLatestRunId(dist) {
    try {
        const res = await fetch(`${GITHUB_PROXY_BASE}?action=list-runs&distributor=${dist}`);
        const data = await res.json();
        if (data.runs && data.runs.length > 0) {
            const recent = data.runs[0];
            // Only use if it was created very recently (within 30 seconds)
            const createdAt = new Date(recent.created_at);
            const now = new Date();
            if (now - createdAt < 30000 && recent.status !== 'completed') {
                return recent.run_id;
            }
        }
    } catch (err) {
        console.error('list-runs fallback error:', err);
    }
    return null;
}

function startWorkflowPolling(dist) {
    stopWorkflowPolling(dist);
    // Poll every 8 seconds
    state.workflowPollingTimers[dist] = setInterval(() => pollWorkflowStatus(dist), 8000);
}

function stopWorkflowPolling(dist) {
    if (state.workflowPollingTimers[dist]) {
        clearInterval(state.workflowPollingTimers[dist]);
        delete state.workflowPollingTimers[dist];
    }
}

async function pollWorkflowStatus(dist) {
    const wf = state.workflowState[dist];
    if (!wf.running) {
        stopWorkflowPolling(dist);
        return;
    }

    try {
        if (wf.runId) {
            const res = await fetch(`${GITHUB_PROXY_BASE}?action=workflow-status&run_id=${wf.runId}`);
            const data = await res.json();

            wf.status = data.status;
            wf.conclusion = data.conclusion;

            // Update running toast text
            const runningToastEl = document.querySelector(`[data-workflow-dist="${dist}"] .workflow-toast-body span`);
            if (runningToastEl) {
                if (data.status === 'queued') runningToastEl.textContent = 'Import queued, waiting to start...';
                else if (data.status === 'in_progress') runningToastEl.textContent = 'Import workflow running...';
            }

            if (data.status === 'completed') {
                wf.running = false;
                stopWorkflowPolling(dist);
                // Remove running toast
                const runningToast = document.querySelector(`[data-workflow-dist="${dist}"]`);
                if (runningToast) runningToast.remove();
                const success = data.conclusion === 'success';
                showWorkflowToast(dist, success, success ? 'Import completed successfully' : `Import ${data.conclusion || 'failed'}`);
                // Store completion time for Last Run display
                state.workflowLastRun[dist] = new Date().toISOString();
                // Re-render stats to update Last Run
                const filterData = state.adminFilterData[dist];
                if (filterData) {
                    if (!filterData.stats) filterData.stats = {};
                    filterData.stats.last_run = state.workflowLastRun[dist];
                    renderMfrStats(dist, filterData);
                }
            }
        } else {
            // Still no run_id — try list-runs fallback
            const foundId = await findLatestRunId(dist);
            if (foundId) {
                wf.runId = foundId;
            }
        }
    } catch (err) {
        console.error('Poll workflow status error:', err);
    }
}

const DIST_LABELS = { tdsynnex: 'TD Synnex', ingram: 'Ingram Micro', adi: 'ADI Global' };

function showWorkflowToast(dist, success, message) {
    const container = document.getElementById('workflowToastContainer');
    if (!container) return;

    const label = DIST_LABELS[dist] || dist;
    const toast = document.createElement('div');
    toast.className = `workflow-toast ${success ? 'workflow-toast--success' : 'workflow-toast--failure'}`;
    toast.innerHTML = `
        <div class="workflow-toast-icon">${success
            ? '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>'
            : '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>'
        }</div>
        <div class="workflow-toast-body">
            <strong>${label}</strong>
            <span>${escHtml(message)}</span>
        </div>
        <button class="workflow-toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    container.appendChild(toast);

    // Auto-dismiss after 10 seconds
    setTimeout(() => {
        if (toast.parentElement) {
            toast.classList.add('workflow-toast--fading');
            setTimeout(() => toast.remove(), 300);
        }
    }, 10000);
}

function showWorkflowRunningToast(dist) {
    const container = document.getElementById('workflowToastContainer');
    if (!container) return;

    // Remove any existing running toast for this distributor
    const existing = container.querySelector(`[data-workflow-dist="${dist}"]`);
    if (existing) existing.remove();

    const label = DIST_LABELS[dist] || dist;
    const toast = document.createElement('div');
    toast.className = 'workflow-toast workflow-toast--running';
    toast.setAttribute('data-workflow-dist', dist);
    toast.innerHTML = `
        <div class="workflow-toast-icon workflow-toast-icon--running">
            <div class="mfr-workflow-spinner-sm"></div>
        </div>
        <div class="workflow-toast-body">
            <strong>${label}</strong>
            <span>Import workflow running...</span>
        </div>
    `;
    container.appendChild(toast);
}

/**
 * Escape HTML entities
 */
function escHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

/**
 * Escape for use in HTML attributes (single-quote safe)
 */
function escAttr(str) {
    return str.replace(/&/g, '&amp;').replace(/'/g, '&#39;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Render the A-Z letter filter bar
 */
function renderLetterBar() {
    const dist = state.adminActiveTab;
    const data = state.adminFilterData[dist];
    const bar = document.getElementById('mfrLetterBar');
    if (!bar) return;

    if (!data) {
        bar.innerHTML = '';
        return;
    }

    const allNames = data.all_known_manufacturers || [];
    // Count which letters have manufacturers
    const letterCounts = {};
    for (const name of allNames) {
        const letter = name.charAt(0).toUpperCase();
        if (/[A-Z]/.test(letter)) {
            letterCounts[letter] = (letterCounts[letter] || 0) + 1;
        }
    }

    const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    const currentFilter = state.adminLetterFilter;

    let html = `<button class="mfr-admin-letter-btn mfr-admin-letter-btn--all${currentFilter === 'All' ? ' active' : ''}" onclick="selectMfrLetter('All')">All</button>`;

    for (const letter of letters) {
        const hasItems = letterCounts[letter] > 0;
        const isActive = currentFilter === letter;
        html += `<button class="mfr-admin-letter-btn${isActive ? ' active' : ''}" ${!hasItems ? 'disabled' : ''} onclick="selectMfrLetter('${letter}')">${letter}</button>`;
    }

    html += `<span class="mfr-letter-scope-divider"></span>`;
    html += `<span class="mfr-letter-scope-group">`;
    html += `<span class="mfr-letter-scope-label">Applies to:</span>`;
    const scope = state.adminLetterScope;
    html += `<button class="mfr-letter-scope-btn${scope === 'available' ? ' active' : ''}" onclick="setMfrLetterScope('available')">Available</button>`;
    html += `<button class="mfr-letter-scope-btn${scope === 'included' ? ' active' : ''}" onclick="setMfrLetterScope('included')">Included</button>`;
    html += `<button class="mfr-letter-scope-btn${scope === 'both' ? ' active' : ''}" onclick="setMfrLetterScope('both')">Both</button>`;
    html += `</span>`;

    bar.innerHTML = html;
}

/**
 * Set letter filter scope (available, included, both)
 */
function setMfrLetterScope(scope) {
    state.adminLetterScope = scope;
    renderLetterBar();
    renderMfrColumns();
}

/**
 * Select a letter filter
 */
function selectMfrLetter(letter) {
    state.adminLetterFilter = letter;
    renderLetterBar();
    renderMfrColumns();
}

// =====================================================
// ZOHO SDK INITIALIZATION
// =====================================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('Widget DOM loaded, initializing...');
    initZohoSDK();
    initEventListeners();
    initDragAndDrop();
    initResize();
    initQueueResize();
    checkProxyStatus();
    updateQueueUI();
    loadIngramManufacturers();  // Pre-populate Ingram manufacturers on init (default distributor)
    loadManufacturerMappings();
    initMfrMappingsResize();
    initBulkPreviewResize();
    initBulkResultsResize();
    initBulkScrollWheelZoom();
    initMfrMappingsTooltip();

    // Set "Group by Manufacturer" checkbox to match default state
    const groupByMfrCheckbox = document.getElementById('groupByMfr');
    if (groupByMfrCheckbox) {
        groupByMfrCheckbox.checked = state.groupByManufacturer;
    }

    // Default: hide discount fields in single mode on initial load
    var initQueuePanel = document.querySelector('.queue-panel');
    var initDiscountToggle = document.getElementById('discountVisibilityToggle');
    var initEyeOpen = document.getElementById('discountEyeOpen');
    var initEyeClosed = document.getElementById('discountEyeClosed');
    if (initQueuePanel) {
        initQueuePanel.classList.add('discount-fields-hidden');
        if (initDiscountToggle) initDiscountToggle.classList.add('fields-hidden');
        if (initEyeOpen) initEyeOpen.style.display = 'none';
        if (initEyeClosed) initEyeClosed.style.display = '';
    }
});

function initZohoSDK() {
    if (typeof ZOHO === 'undefined') {
        console.warn('ZOHO SDK not loaded. Running in standalone mode.');
        showStatus('Running in standalone mode (Zoho SDK not available)', 'info');
        return;
    }

    ZOHO.embeddedApp.init();
    console.log('ZOHO.embeddedApp.init() called');

    ZOHO.embeddedApp.on("PageLoad", function(data) {
        console.log('PageLoad event received:', data);
        state.parentContext = data;

        // Store pre-fetched manufacturers from Client Script (Phase 3)
        if (data && data.manufacturers && Array.isArray(data.manufacturers)) {
            state.prefetchedManufacturers = data.manufacturers.map(m => {
                // Handle both {id, name} objects and plain strings
                return typeof m === 'string' ? m : (m.name || m.Name || '');
            }).filter(name => name.length > 0).sort();
            console.log(`[MfrResolution] Received ${state.prefetchedManufacturers.length} manufacturers from Zoho`);
        }

        showStatus('Widget loaded. Select a manufacturer or enter a SKU to begin.', 'info');
    });

    ZOHO.embeddedApp.on("NotifyAndWait", function(data) {
        console.log('NotifyAndWait event received:', data);
        state.pendingResponseId = data.id;
        state.parentContext = data.data || {};

        // Also check NotifyAndWait for manufacturers (might come later)
        const eventData = data.data || {};
        if (eventData.manufacturers && Array.isArray(eventData.manufacturers) && state.prefetchedManufacturers.length === 0) {
            state.prefetchedManufacturers = eventData.manufacturers.map(m => {
                return typeof m === 'string' ? m : (m.name || m.Name || '');
            }).filter(name => name.length > 0).sort();
            console.log(`[MfrResolution] Received ${state.prefetchedManufacturers.length} manufacturers from NotifyAndWait`);
        }

        showStatus('Ready to search. Select products and click "Add to Queue".', 'info');
    });
}

// =====================================================
// EVENT LISTENERS
// =====================================================
function initEventListeners() {
    const mfrSearch = document.getElementById('manufacturerSearch');
    if (mfrSearch) {
        mfrSearch.addEventListener('input', debounceManufacturerSearch);
    }

    // SKU search field (single field for both SKU-first and filter modes)
    const skuSearch = document.getElementById('skuSearch');
    if (skuSearch) {
        skuSearch.addEventListener('input', () => {
            state.skuKeyword = skuSearch.value.trim();
        });
        skuSearch.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                handleLoadProducts();
            }
        });
    }

    const selectAll = document.getElementById('selectAll');
    if (selectAll) {
        selectAll.addEventListener('change', toggleSelectAll);
    }
}

// =====================================================
// RESIZE FUNCTIONALITY
// =====================================================
function initResize() {
    const resizeHandle = document.getElementById('resizeHandle');
    const tableContainer = document.querySelector('.table-container');

    if (!resizeHandle || !tableContainer) return;

    resizeHandle.addEventListener('mousedown', (e) => {
        isResizing = true;
        resizeStartY = e.clientY;
        resizeStartHeight = tableContainer.offsetHeight;
        document.body.style.cursor = 'ns-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;

        const deltaY = e.clientY - resizeStartY;
        const newHeight = Math.max(100, Math.min(1250, resizeStartHeight + deltaY));
        tableContainer.style.maxHeight = newHeight + 'px';
    });

    document.addEventListener('mouseup', () => {
        if (isResizing) {
            isResizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// =====================================================
// QUEUE PANEL HORIZONTAL RESIZE
// =====================================================
function initQueueResize() {
    const resizeHandle = document.getElementById('queueResizeHandle');
    const rightPanel = document.getElementById('rightPanel');

    if (!resizeHandle || !rightPanel) return;

    resizeHandle.addEventListener('mousedown', (e) => {
        isResizingQueue = true;
        queueResizeStartX = e.clientX;
        queueResizeStartWidth = rightPanel.offsetWidth;
        document.body.style.cursor = 'ew-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizingQueue) return;

        // Dragging left increases width, dragging right decreases
        const deltaX = queueResizeStartX - e.clientX;
        const newWidth = Math.max(240, Math.min(600, queueResizeStartWidth + deltaX));
        rightPanel.style.width = newWidth + 'px';

        // Toggle narrow class for responsive stacking
        if (newWidth < 300) {
            rightPanel.classList.add('narrow');
        } else {
            rightPanel.classList.remove('narrow');
        }
    });

    document.addEventListener('mouseup', () => {
        if (isResizingQueue) {
            isResizingQueue = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// =====================================================
// TOOLTIP POSITIONING (Fixed Position + JS)
// =====================================================
/**
 * Initialize tooltip positioning for info buttons
 * Uses position: fixed and calculates position on hover
 * Call this after dynamically rendering tooltip elements
 */
function initMfrResolutionTooltips() {
    document.querySelectorAll('.mfr-resolution-table .tooltip-wrapper').forEach(wrapper => {
        const btn = wrapper.querySelector('.info-btn');
        const tooltip = wrapper.querySelector('.tooltip');

        if (!btn || !tooltip) return;

        // Remove any existing listeners (avoid duplicates)
        const newBtn = btn.cloneNode(true);
        btn.parentNode.replaceChild(newBtn, btn);

        newBtn.addEventListener('mouseenter', (e) => {
            const rect = newBtn.getBoundingClientRect();
            // Center the 220px tooltip above the button
            tooltip.style.left = (rect.left + rect.width/2 - 110) + 'px';
            tooltip.style.top = (rect.top - tooltip.offsetHeight - 8) + 'px';
            tooltip.classList.add('visible');
        });

        newBtn.addEventListener('mouseleave', () => {
            tooltip.classList.remove('visible');
        });
    });
}

/**
 * Initialize tooltip for the manufacturer mappings button
 * Uses same pattern as mfr resolution tooltips
 */
function initMfrMappingsTooltip() {
    const wrapper = document.querySelector('.mfr-mappings-tooltip-wrapper');
    if (!wrapper) return;

    const btn = wrapper.querySelector('.mfr-mappings-btn');
    const tooltip = wrapper.querySelector('.tooltip');

    if (!btn || !tooltip) return;

    btn.addEventListener('mouseenter', (e) => {
        const rect = btn.getBoundingClientRect();
        // Position tooltip below the button (since it's in the header area)
        tooltip.style.left = (rect.left + rect.width/2 - 110) + 'px';
        tooltip.style.top = (rect.bottom + 8) + 'px';
        tooltip.classList.add('visible');
    });

    btn.addEventListener('mouseleave', () => {
        tooltip.classList.remove('visible');
    });

    // Also init bulk mode tooltip
    const bulkWrapper = document.querySelector('.bulk-mfr-mappings-tooltip-wrapper');
    if (bulkWrapper) {
        const bulkBtnEl = bulkWrapper.querySelector('.mfr-mappings-btn');
        const bulkTooltip = bulkWrapper.querySelector('.tooltip');

        if (bulkBtnEl && bulkTooltip) {
            bulkBtnEl.addEventListener('mouseenter', (e) => {
                const rect = bulkBtnEl.getBoundingClientRect();
                bulkTooltip.style.left = (rect.left + rect.width/2 - 110) + 'px';
                bulkTooltip.style.top = (rect.bottom + 8) + 'px';
                bulkTooltip.classList.add('visible');
            });

            bulkBtnEl.addEventListener('mouseleave', () => {
                bulkTooltip.classList.remove('visible');
            });
        }
    }
}

// =====================================================
// DRAG AND DROP FOR QUEUE
// =====================================================
function initDragAndDrop() {
    const queueItems = document.getElementById('queueItems');
    if (!queueItems) return;

    queueItems.addEventListener('dragstart', handleDragStart);
    queueItems.addEventListener('dragend', handleDragEnd);
    queueItems.addEventListener('dragover', handleDragOver);
    queueItems.addEventListener('drop', handleDrop);
}

function handleDragStart(e) {
    // Handle manufacturer group dragging
    if (e.target.classList.contains('queue-mfr-group')) {
        draggedGroup = e.target;
        e.target.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'group:' + e.target.dataset.manufacturer);
        return;
    }

    if (!e.target.classList.contains('queue-item')) return;

    draggedItem = e.target;
    e.target.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', e.target.dataset.partNumber);
}

function handleDragEnd(e) {
    if (draggedItem) {
        draggedItem.classList.remove('dragging');
        draggedItem = null;
    }
    if (draggedGroup) {
        draggedGroup.classList.remove('dragging');
        draggedGroup = null;
    }
    document.querySelectorAll('.queue-item, .queue-mfr-group').forEach(item => {
        item.classList.remove('drag-over');
    });
}

function handleDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';

    const queueItems = document.getElementById('queueItems');

    if (draggedGroup && state.groupByManufacturer) {
        // Handle group reordering
        const afterGroup = getDragAfterGroup(e.clientY);
        const groupWithItems = getGroupElements(draggedGroup.dataset.manufacturer);

        if (afterGroup == null) {
            // Move to end
            groupWithItems.forEach(el => queueItems.appendChild(el));
        } else {
            // Move before the target group
            const beforeEl = afterGroup;
            groupWithItems.forEach(el => queueItems.insertBefore(el, beforeEl));
        }
    } else if (draggedItem) {
        const afterElement = getDragAfterElement(e.clientY);

        if (afterElement == null) {
            queueItems.appendChild(draggedItem);
        } else {
            queueItems.insertBefore(draggedItem, afterElement);
        }
    }
}

function handleDrop(e) {
    e.preventDefault();

    if (state.groupByManufacturer && draggedGroup) {
        // Reorder by manufacturer groups
        reorderByGroups();
    } else {
        // Reorder queue based on new DOM order
        const newOrder = [];
        document.querySelectorAll('.queue-item').forEach(item => {
            const partNumber = item.dataset.partNumber;
            const product = getActiveQueue().find(p =>
                (p.ingramPartNumber || p.vendorPartNumber) === partNumber
            );
            if (product) {
                newOrder.push(product);
            }
        });

        setActiveQueue(newOrder);
        console.log('[Queue] Reordered:', getActiveQueue().map(p => p.vendorPartNumber));
    }
}

function getDragAfterElement(y) {
    const draggableElements = [...document.querySelectorAll('.queue-item:not(.dragging)')];

    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;

        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function getDragAfterGroup(y) {
    const groupElements = [...document.querySelectorAll('.queue-mfr-group:not(.dragging)')];

    return groupElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;

        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

function getGroupElements(manufacturer) {
    const elements = [];
    const groupHeader = document.querySelector(`.queue-mfr-group[data-manufacturer="${manufacturer}"]`);
    if (groupHeader) {
        elements.push(groupHeader);
        // Get all items after this header until the next header
        let sibling = groupHeader.nextElementSibling;
        while (sibling && !sibling.classList.contains('queue-mfr-group')) {
            elements.push(sibling);
            sibling = sibling.nextElementSibling;
        }
    }
    return elements;
}

function reorderByGroups() {
    // Get the new order of manufacturers from the DOM
    const mfrOrder = [];
    document.querySelectorAll('.queue-mfr-group').forEach(group => {
        mfrOrder.push(group.dataset.manufacturer);
    });

    // Reorder queuedProducts based on manufacturer order
    const newOrder = [];
    mfrOrder.forEach(mfr => {
        getActiveQueue()
            .filter(p => (p.vendorName || p.manufacturer || 'Unknown') === mfr)
            .forEach(p => newOrder.push(p));
    });

    setActiveQueue(newOrder);
    console.log('[Queue] Reordered by groups:', mfrOrder);
}

// =====================================================
// SCROLL-TO-FOCUS HELPER
// =====================================================
function scrollToPanel(el) {
    if (!el) return;
    if (typeof el === 'string') el = document.getElementById(el);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// =====================================================
// RAW API TOGGLE
// =====================================================
function toggleRawApi() {
    state.rawApiVisible = !state.rawApiVisible;
    const container = document.getElementById('rawApiContainer');
    const toggle = document.getElementById('rawApiToggle');

    if (container && toggle) {
        container.style.display = state.rawApiVisible ? 'block' : 'none';
        toggle.classList.toggle('active', state.rawApiVisible);
        if (state.rawApiVisible) scrollToPanel(container);
    }
}

// =====================================================
// GROUP BY MANUFACTURER TOGGLE
// =====================================================
function toggleGroupByManufacturer() {
    const checkbox = document.getElementById('groupByMfr');
    state.groupByManufacturer = checkbox ? checkbox.checked : false;
    renderQueueItems();
}

function setPricingMode(mode) {
    if (state.searchMode === 'bulk') {
        bulkState.pricingMode = mode;
    } else {
        state.pricingMode = mode;
    }
    document.querySelectorAll('#pricingToggle .pricing-toggle-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.price === mode));
    updateQueueUI();
    updateQueueTotals();
}

// =====================================================
// DISTRIBUTOR SELECTION
// =====================================================
function selectDistributor(distributor) {
    if (DISTRIBUTORS[distributor]?.disabled) {
        showStatus(`${DISTRIBUTORS[distributor].name} integration coming soon`, 'info');
        return;
    }

    state.currentDistributor = distributor;

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.distributor === distributor);
    });

    // Clear product details when switching distributors (queue is preserved)
    hideProductDetails();

    // Show/hide distributor-specific filters
    const cat3Field = document.getElementById('cat3FilterField');
    const skuTypeField = document.getElementById('skuTypeFilterField');

    // Remove/add Ingram mode from manufacturer combo
    const mfrComboEl = document.querySelector('.mfr-combo');
    if (mfrComboEl) {
        mfrComboEl.classList.remove('ingram-mode', 'adi-mode', 'tdsynnex-mode');
    }

    if (distributor === 'tdsynnex') {
        // TD Synnex: Show cat3, hide SKU type
        // Restore subcategory filter (may have been hidden by ADI)
        const subCatFieldRestore = document.getElementById('subcategorySelect');
        if (subCatFieldRestore) subCatFieldRestore.closest('.filter-field').style.display = '';
        if (cat3Field) cat3Field.style.display = '';
        if (skuTypeField) skuTypeField.style.display = 'none';
        // TD Synnex: pre-populate manufacturer dropdown, hide search input
        if (mfrComboEl) mfrComboEl.classList.add('tdsynnex-mode');
        loadTDSynnexManufacturers();
    } else if (distributor === 'adi') {
        // ADI Global: Hide cat3, hide SKU type, hide subcategory
        if (cat3Field) cat3Field.style.display = 'none';
        if (skuTypeField) skuTypeField.style.display = 'none';
        // Hide subcategory filter for ADI
        const subCatField = document.getElementById('subcategorySelect');
        if (subCatField) subCatField.closest('.filter-field').style.display = 'none';
        // ADI: pre-populate manufacturer dropdown, hide search input
        if (mfrComboEl) mfrComboEl.classList.add('adi-mode');
        // Pre-populate manufacturer and category_1 dropdowns
        loadADIManufacturers();
    } else {
        // Ingram: Hide cat3, show Media Type filter (DB-driven)
        // Restore subcategory filter (may have been hidden by ADI)
        const subCatFieldRestore = document.getElementById('subcategorySelect');
        if (subCatFieldRestore) subCatFieldRestore.closest('.filter-field').style.display = '';
        if (cat3Field) cat3Field.style.display = 'none';
        if (skuTypeField) skuTypeField.style.display = '';
        // Ingram: pre-populate manufacturer dropdown, hide search input
        if (mfrComboEl) mfrComboEl.classList.add('ingram-mode');
        // Ingram uses DB media_type codes — update label and reset to dynamic dropdown
        const skuTypeLabel = document.getElementById('skuTypeLabel');
        if (skuTypeLabel) {
            // Update label text without destroying the count badge span
            const countSpan = skuTypeLabel.querySelector('.count-badge');
            if (countSpan) {
                // Preserve the span, just update the text before it
                skuTypeLabel.firstChild.textContent = 'Media Type ';
            } else {
                skuTypeLabel.innerHTML = 'Media Type <span class="count-badge" id="skuTypeCount"></span>';
            }
        }
        const skuTypeSelect = document.getElementById('skuTypeSelect');
        if (skuTypeSelect) skuTypeSelect.innerHTML = '<option value="">-- Any --</option>';
    }

    resetFilters();
    if (state.searchMode === 'bulk') {
        showStatus(`Switched to ${DISTRIBUTORS[distributor].name}. Upload a distributor quote or paste SKUs to be parsed and loaded.`, 'info');
    } else {
        showStatus(`Switched to ${DISTRIBUTORS[distributor].name}. Search by manufacturer or SKU.`, 'info');
    }

    // Update bulk distributor badges if in bulk mode
    if (state.searchMode === 'bulk') {
        updateBulkDistributorBadges();
        // Re-parse paste content for new distributor's format
        var bulkSkuInput = document.getElementById('bulkPasteArea');
        if (bulkSkuInput && bulkSkuInput.value.trim()) {
            bulkParsePastedSKUs();
        }
        bulkUpdateLoadButtonState();
    }
}

function updateBulkDistributorBadges() {
    const dist = DISTRIBUTORS[state.currentDistributor];
    const name = dist ? dist.name : '';
    const badge = document.getElementById('bulkDistributorBadge');
    if (badge) badge.textContent = name;
}

// =====================================================
// PROXY STATUS CHECK
// =====================================================
async function checkProxyStatus() {
    const indicator = document.getElementById('statusIndicator');
    const statusText = document.getElementById('statusText');

    try {
        const response = await fetch(`${PROXY_BASE}?action=status`);
        const data = await response.json();

        if (data.authenticated) {
            indicator.classList.add('connected');
            statusText.textContent = 'Connected';
            state.isAuthenticated = true;
        } else if (data.configured) {
            statusText.textContent = 'Not authenticated';
            await authenticate();
        } else {
            statusText.textContent = 'Not configured';
            showStatus('Proxy server not configured. Check credentials.', 'error');
        }
    } catch (error) {
        indicator.classList.remove('connected');
        statusText.textContent = 'Offline';
        showStatus('Cannot connect to proxy server.', 'error');
    }
}

async function authenticate() {
    try {
        const response = await fetch(`${PROXY_BASE}?action=auth`);
        const data = await response.json();

        if (data.success) {
            document.getElementById('statusIndicator').classList.add('connected');
            document.getElementById('statusText').textContent = 'Connected';
            state.isAuthenticated = true;
            showStatus('Authentication successful. Search for a manufacturer.', 'success');
        }
    } catch (error) {
        showStatus('Authentication failed: ' + error.message, 'error');
    }
}

// =====================================================
// TD SYNNEX API FUNCTIONS
// =====================================================

// TD Synnex warehouse location mapping for display
const TDSYNNEX_WAREHOUSES = {
    qty_miami_fl: { id: 'MIA', location: 'Miami, FL' },
    qty_tracy_ca: { id: 'TRC', location: 'Tracy, CA' },
    qty_romeoville_il: { id: 'ROM', location: 'Romeoville, IL' },
    qty_southaven_ms: { id: 'SHV', location: 'Southaven, MS' },
    qty_columbus_oh: { id: 'COL', location: 'Columbus, OH' },
    qty_suwanee_ga: { id: 'SUW', location: 'Suwanee, GA' },
    qty_chino_ca: { id: 'CHI', location: 'Chino, CA' },
    qty_swedesboro_nj: { id: 'SWD', location: 'Swedesboro, NJ' },
    qty_south_bend_in: { id: 'SBI', location: 'South Bend, IN' },
    qty_fort_worth_tx: { id: 'FTW', location: 'Fort Worth, TX' },
    qty_fontana_ca: { id: 'FON', location: 'Fontana, CA' }
};

// TD SYNNEX Kit/Standalone formatter
function formatKitStandalone(flag) {
    if (!flag) return '-';
    return flag.toUpperCase() === 'K' ? 'Kit' : 'Standalone';
}

// TD SYNNEX ABC Code formatter
function formatABCCode(code) {
    const defs = {
        'A': 'Active',
        'B': 'Special Order',
        'C': 'EOL',
        'T': 'To Be Discontinued'
    };
    if (!code) return '-';
    return defs[code.toUpperCase()] || code;
}

// Check if TD SYNNEX product is "New" (created within 90 days)
function isTDSynnexNew(createdDate) {
    if (!createdDate) return false;
    const created = new Date(createdDate);
    const ninetyDaysAgo = new Date();
    ninetyDaysAgo.setDate(ninetyDaysAgo.getDate() - 90);
    return created >= ninetyDaysAgo;
}

// Check if TD SYNNEX product is Licensed
function isTDSynnexLicensed(assignedUse, longDescription) {
    const searchTerm = 'licens';
    const assignedMatch = assignedUse && assignedUse.toLowerCase().includes(searchTerm);
    const descMatch = longDescription && longDescription.toLowerCase().includes(searchTerm);
    return assignedMatch || descMatch;
}

// Check if TD SYNNEX product is a Service SKU
function isTDSynnexServiceSku(cat1, cat2) {
    const cat1Lower = (cat1 || '').toLowerCase();
    const cat2Lower = (cat2 || '').toLowerCase();
    return cat1Lower === 'service / support' ||
           cat2Lower.includes('service') ||
           cat2Lower.includes('support');
}

// Format YYMMDD date to MM/DD/YYYY
function formatPromoDate(yymmdd) {
    if (!yymmdd || yymmdd.length !== 6) return yymmdd || 'N/A';
    const yy = yymmdd.substring(0, 2);
    const mm = yymmdd.substring(2, 4);
    const dd = yymmdd.substring(4, 6);
    return `${mm}/${dd}/20${yy}`;
}

// Derive SKU Type from physical dimensions
function deriveSKUType(weight, length, width, height) {
    // If all dimensions are 0 or null, it's Digital
    const allZero = (weight || 0) === 0 &&
                    (length || 0) === 0 &&
                    (width || 0) === 0 &&
                    (height || 0) === 0;
    return allZero ? 'Digital' : 'Physical';
}

// Fetch warehouse availability from TD SYNNEX XML API
async function fetchTDSynnexWarehouseAvailability(synnexSKU) {
    if (!synnexSKU) return { warehouses: [], totalQty: 0, totalOnOrder: 0, status: null };
    try {
        const response = await fetch(
            `${TDSYNNEX_PROXY_BASE}?action=availability&synnexSKU=${encodeURIComponent(synnexSKU)}`
        );
        const data = await response.json();
        return {
            warehouses: data.warehouses || [],
            totalQty: data.totalQty || 0,
            totalOnOrder: data.totalOnOrder || 0,
            status: data.status || null
        };
    } catch (error) {
        console.error('[TD SYNNEX] Warehouse fetch error:', error);
        return { warehouses: [], totalQty: 0, totalOnOrder: 0, status: null };
    }
}

async function searchTDSynnexManufacturers(searchTerm) {
    const response = await fetch(
        `${SUPABASE_URL}/rest/v1/rpc/get_tdsynnex_manufacturers`,
        {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify({ search_term: searchTerm, result_limit: 100 })
        }
    );
    const data = await response.json();
    return (data || []).map(m => m.manufacturer_name);
}

async function loadIngramManufacturers() {
    const select = document.getElementById('manufacturerSelect');
    const countEl = document.getElementById('mfrCount');
    try {
        const response = await fetch(
            `${SUPABASE_URL}/rest/v1/rpc/get_ingram_manufacturers`,
            {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                },
                body: JSON.stringify({ p_search: null })
            }
        );
        if (!response.ok) throw new Error(`Failed: ${response.status}`);
        const data = await response.json();
        const manufacturers = data.map(r => r.manufacturer).filter(Boolean).sort();

        select.innerHTML = '<option value="">-- Select Manufacturer --</option>' +
            manufacturers.map(m => `<option value="${m}">${m}</option>`).join('');
        if (countEl) countEl.textContent = `(${manufacturers.length})`;
        select.disabled = false;
    } catch (error) {
        console.error('[Ingram] Failed to load manufacturers:', error);
        select.innerHTML = '<option value="">Error loading manufacturers</option>';
    }
}

async function loadADIManufacturers() {
    const select = document.getElementById('manufacturerSelect');
    const countEl = document.getElementById('mfrCount');
    try {
        const response = await fetch(
            `${SUPABASE_URL}/rest/v1/rpc/get_adi_manufacturers`,
            {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                },
                body: JSON.stringify({})
            }
        );
        if (!response.ok) throw new Error(`Failed: ${response.status}`);
        const data = await response.json();
        const manufacturers = data.map(r => r.manufacturer).filter(Boolean).sort();

        select.innerHTML = '<option value="">-- Select Manufacturer --</option>' +
            manufacturers.map(m => `<option value="${m}">${m}</option>`).join('');
        if (countEl) countEl.textContent = `(${manufacturers.length})`;
        select.disabled = false;
    } catch (error) {
        console.error('[ADI] Failed to load manufacturers:', error);
        select.innerHTML = '<option value="">Error loading manufacturers</option>';
    }
}

async function loadTDSynnexManufacturers() {
    const select = document.getElementById('manufacturerSelect');
    const countEl = document.getElementById('mfrCount');
    try {
        const response = await fetch(
            `${SUPABASE_URL}/rest/v1/rpc/get_tdsynnex_manufacturers`,
            {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                },
                body: JSON.stringify({})
            }
        );
        if (!response.ok) throw new Error(`Failed: ${response.status}`);
        const data = await response.json();
        const manufacturers = data.map(r => r.manufacturer_name).filter(Boolean).sort();

        select.innerHTML = '<option value="">-- Select Manufacturer --</option>' +
            manufacturers.map(m => `<option value="${m}">${m}</option>`).join('');
        if (countEl) countEl.textContent = `(${manufacturers.length})`;
        select.disabled = false;
    } catch (error) {
        console.error('[TD Synnex] Failed to load manufacturers:', error);
        select.innerHTML = '<option value="">Error loading manufacturers</option>';
    }
}

// Lazy API verification for Ingram manufacturers
// Silently verifies unverified manufacturer names against Ingram catalog API
// Fire-and-forget — never blocks UI or shows errors to user
async function verifyIngramManufacturers(products) {
    try {
        // Only process Ingram products
        const ingramProducts = products.filter(p => p._source === 'ingram');
        if (ingramProducts.length === 0) return;

        // Extract unique manufacturer names, skip already-verified this session
        const manufacturers = [...new Set(ingramProducts.map(p => p.vendorName).filter(Boolean))];
        const unchecked = manufacturers.filter(m => !state.verifiedIngramMfrs.has(m));
        if (unchecked.length === 0) return;

        // Ask Supabase which of these are actually unverified in DB
        const unverifiedResp = await fetch(
            `${SUPABASE_URL}/rest/v1/rpc/get_unverified_ingram_manufacturers`,
            {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                },
                body: JSON.stringify({ p_manufacturers: unchecked })
            }
        );
        if (!unverifiedResp.ok) return;
        const unverifiedList = await unverifiedResp.json();
        const unverifiedNames = new Set((unverifiedList || []).map(r => r.manufacturer));

        // Mark all checked manufacturers as verified this session (even if already verified in DB)
        unchecked.forEach(m => state.verifiedIngramMfrs.add(m));

        if (unverifiedNames.size === 0) return;

        // For each unverified manufacturer, pick first product with a VPN and verify via API
        for (const mfrName of unverifiedNames) {
            try {
                const sample = ingramProducts.find(p => p.vendorName === mfrName && p.vendorPartNumber);
                if (!sample) continue;

                // Call Ingram catalog API with VPN as keyword (productsWithPricing allows keyword-only, no vendor required)
                const skuResp = await fetch(
                    `${PROXY_BASE}?action=productsWithPricing&keyword=${encodeURIComponent(sample.vendorPartNumber)}`,
                    { headers: { 'Accept': 'application/json' } }
                );
                if (!skuResp.ok) continue;
                const skuData = await skuResp.json();

                // Find matching product in API response by VPN
                const apiProducts = skuData.products || [];
                const match = apiProducts.find(p =>
                    (p.vendorPartNumber || '').toUpperCase() === sample.vendorPartNumber.toUpperCase()
                );
                if (!match || !match.vendorName) continue;

                const apiMfrName = match.vendorName.trim();
                if (!apiMfrName) continue;

                // Update DB: set api_verified=true, update manufacturer name if different
                await fetch(
                    `${SUPABASE_URL}/rest/v1/rpc/verify_ingram_manufacturer`,
                    {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'apikey': SUPABASE_ANON_KEY,
                            'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                        },
                        body: JSON.stringify({
                            p_manufacturer: mfrName,
                            p_api_manufacturer_name: apiMfrName
                        })
                    }
                );
                console.log(`[Ingram Verify] ${mfrName} → ${apiMfrName} (${mfrName === apiMfrName ? 'confirmed' : 'updated'})`);
            } catch (e) {
                // Silent failure per manufacturer — don't stop verifying others
                console.warn(`[Ingram Verify] Failed for ${mfrName}:`, e.message);
            }
        }
    } catch (e) {
        // Silent failure — verification is best-effort
        console.warn('[Ingram Verify] Verification failed:', e.message);
    }
}

async function loadTDSynnexCategories(manufacturer, cat1 = null, cat2 = null) {
    // Determine which level RPC to call based on params
    let rpcName, body, descField;
    if (!cat1) {
        rpcName = 'get_tdsynnex_categories_level1';
        body = { p_manufacturer: manufacturer };
        descField = 'cat_description_1';
    } else if (!cat2) {
        rpcName = 'get_tdsynnex_categories_level2';
        body = { p_manufacturer: manufacturer, p_cat1: cat1 };
        descField = 'cat_description_2';
    } else {
        rpcName = 'get_tdsynnex_categories_level3';
        body = { p_manufacturer: manufacturer, p_cat1: cat1, p_cat2: cat2 };
        descField = 'cat_description_3';
    }

    const response = await fetch(
        `${SUPABASE_URL}/rest/v1/rpc/${rpcName}`,
        {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify(body)
        }
    );
    const data = await response.json();
    return (data || []).map(c => ({ name: c[descField], count: c.product_count }));
}

async function searchTDSynnexProducts(manufacturer, options = {}) {
    const { search = '', cat1 = '', cat2 = '', cat3 = '', limit = PAGE_SIZE, offset = 0 } = options;

    // Build RPC params — use the overload matching the category depth
    const countBody = {
        p_manufacturer: manufacturer,
        p_search: search || null
    };
    if (cat1) countBody.p_cat1 = cat1;
    if (cat1) countBody.p_cat2 = cat2 || null;
    if (cat2) countBody.p_cat3 = cat3 || null;

    // Products body adds pagination params
    const rpcBody = { ...countBody, p_limit: limit, p_offset: offset };

    // Fetch products and count in parallel
    const [productsResponse, countResponse] = await Promise.all([
        fetch(`${SUPABASE_URL}/rest/v1/rpc/search_tdsynnex_products`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify(rpcBody)
        }),
        fetch(`${SUPABASE_URL}/rest/v1/rpc/search_tdsynnex_products_count`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify(countBody)
        })
    ]);

    const products = await productsResponse.json();
    const totalCount = await countResponse.json();

    if (Array.isArray(products)) {
        return {
            products: products.map(mapTDSynnexProduct),
            totalCount: totalCount || 0,
            pagination: {
                page: Math.floor(offset / limit) + 1,
                pageSize: limit,
                totalPages: Math.ceil((totalCount || 0) / limit),
                totalRecords: totalCount || 0
            }
        };
    }
    return { products: [], totalCount: 0, pagination: null };
}

// Map TD Synnex product to widget's expected format (matching Ingram structure)
function mapTDSynnexProduct(product) {
    // Derive SKU Type from physical dimensions
    const skuType = deriveSKUType(
        product.ship_weight,
        product.length,
        product.width,
        product.height
    );

    return {
        // Core identifiers
        vendorPartNumber: product.manufacturer_part_number,
        ingramPartNumber: product.manufacturer_part_number, // Use MPN for display consistency
        distributorPartNumber: product.td_synnex_sku,
        tdSynnexSkuNumber: product.td_synnex_sku_number, // Numeric SKU for API calls (Field 5)

        // Product info
        description: product.part_description,
        vendorName: product.manufacturer_name,
        category: product.cat_description_1,
        subCategory: product.cat_description_2,
        cat3: product.cat_description_3,

        // Extended description
        extraDescription: product.long_description || product.long_description_1 || '',

        // Pricing (TD Synnex has msrp directly)
        retailPrice: product.msrp,
        unitCost: product.unit_cost,
        contractPrice: product.contract_price,
        pricingData: {
            pricing: {
                retailPrice: product.msrp,
                customerPrice: product.contract_price || product.unit_cost
            },
            availability: buildTDSynnexAvailability(product)
        },

        // Flags and derived fields
        productType: product.kit_standalone_flag || '',
        kitStandaloneFlag: product.kit_standalone_flag,
        type: skuType === 'Digital' ? 'TS::digital' : 'TS::physical',
        skuType: skuType,
        abcCode: product.abc_code,
        skuAttributes: product.sku_attributes,
        replacementSku: product.replacement_sku,

        // Derived boolean flags
        isLicensed: isTDSynnexLicensed(product.td_assigned_use, product.long_description),
        isServiceSku: isTDSynnexServiceSku(product.cat_description_1, product.cat_description_2),
        isNew: isTDSynnexNew(product.sku_created_date),
        isDigital: skuType === 'Digital',
        isDiscontinued: product.abc_code === 'C' || product.abc_code === 'T',

        // Physical dimensions (for reference)
        shipWeight: product.ship_weight,
        length: product.length,
        width: product.width,
        height: product.height,

        // Additional data
        upcCode: product.upc_code,
        commodityName: product.commodity_name,
        lastUpdated: product.last_updated,
        skuCreatedDate: product.sku_created_date,
        tdAssignedUse: product.td_assigned_use,

        // Promo info
        promoFlag: product.promo_flag,
        promoComment: product.promo_comment,
        promoExpiration: product.promo_expiration,
        etaDate: product.eta_date,

        // Source marker
        _source: 'tdsynnex',
        _rawProduct: product
    };
}

function mapADIGlobalProduct(row) {
    return {
        // Core identifiers — MPN is product_code_mpn, VPN/SKU is item
        vendorPartNumber: row.product_code_mpn || '',
        ingramPartNumber: row.product_code_mpn || '', // Use MPN for display consistency
        distributorPartNumber: row.item || '',
        adiSku: row.item || '',

        // Product info
        description: row.item_desc || '',
        vendorName: row.manufacturer || '',
        category: row.category_1 || '',
        category2: row.category_2 || '',

        // Extended description — use vendor_part_desc if available
        extraDescription: row.vendor_part_desc || row.item_desc || '',

        // Pricing — ADI has current_price (reseller) and msrp
        retailPrice: row.msrp ? parseFloat(row.msrp) : null,
        pricingData: {
            _dbSource: true,
            pricing: {
                retailPrice: row.msrp ? parseFloat(row.msrp) : null,
                customerPrice: row.current_price ? parseFloat(row.current_price) : null
            }
        },
        resellerPrice: row.current_price ? parseFloat(row.current_price) : null,

        // UPC
        upcCode: row.upc_number || '',

        // Source marker
        _source: 'adi',
        _rawProduct: row
    };
}

// Build availability data in Ingram-compatible format
function buildTDSynnexAvailability(product) {
    const qtyTotal = product.qty_total ?? 0;
    const isVirtual = qtyTotal === 9999;

    // Build warehouse breakdown
    const availabilityByWarehouse = [];
    for (const [field, info] of Object.entries(TDSYNNEX_WAREHOUSES)) {
        const qty = product[field];
        if (qty !== null && qty !== undefined && qty > 0) {
            availabilityByWarehouse.push({
                warehouseId: info.id,
                location: info.location,
                quantityAvailable: qty,
                quantityBackordered: 0
            });
        }
    }

    return {
        available: qtyTotal > 0,
        totalAvailability: isVirtual ? 'Unlimited' : qtyTotal,
        availabilityByWarehouse: availabilityByWarehouse,
        isVirtual: isVirtual
    };
}

// =====================================================
// =====================================================
// MANUFACTURER SEARCH
// =====================================================
function debounceManufacturerSearch() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(searchManufacturers, 300);
}

async function searchManufacturers() {
    if (state.currentDistributor === 'ingram') return;
    const searchTerm = document.getElementById('manufacturerSearch').value.trim();
    const select = document.getElementById('manufacturerSelect');

    if (searchTerm.length < 2) {
        select.innerHTML = '<option value="">Type 2+ characters to search...</option>';
        document.getElementById('mfrCount').textContent = '';
        return;
    }

    showStatus(`Searching manufacturers matching "${searchTerm}"...`, 'loading');

    try {
        let manufacturers = [];

        if (state.currentDistributor === 'tdsynnex') {
            // TD Synnex: Use dedicated edge function
            manufacturers = await searchTDSynnexManufacturers(searchTerm);
        } else {
            // Ingram: Query Supabase DB (normalized manufacturer names)
            const response = await fetch(
                `${SUPABASE_URL}/rest/v1/rpc/get_ingram_manufacturers`,
                {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                    },
                    body: JSON.stringify({ p_search: searchTerm })
                }
            );
            const data = await response.json();
            manufacturers = (data || []).map(row => row.manufacturer);
        }

        select.innerHTML = '<option value="">-- Select a manufacturer --</option>';

        if (manufacturers.length > 0) {
            manufacturers.forEach(mfr => {
                const option = document.createElement('option');
                option.value = mfr;
                option.textContent = mfr;
                select.appendChild(option);
            });
            document.getElementById('mfrCount').textContent = `(${manufacturers.length})`;
            showStatus(`Found ${manufacturers.length} manufacturers`, 'success');
        } else {
            select.innerHTML = '<option value="">No manufacturers found</option>';
            document.getElementById('mfrCount').textContent = '(0)';
            showStatus('No manufacturers found. Try a different search term.', 'info');
        }
    } catch (error) {
        showStatus('Error searching: ' + error.message, 'error');
    }
}

// =====================================================
// SKU-FIRST SEARCH FUNCTIONS
// =====================================================

/**
 * Handle the Load Products button click
 * Determines whether to use manufacturer-first or SKU-first flow
 */
function handleLoadProducts() {
    const skuValue = document.getElementById('skuSearch')?.value.trim() || '';

    if (state.manufacturer) {
        // Manufacturer is selected - use normal flow
        loadProducts(1);
    } else if (skuValue.length >= 2) {
        // No manufacturer but have SKU - use SKU-first flow
        lookupManufacturersFromSKU(skuValue);
    } else {
        // Neither manufacturer nor valid SKU
        showStatus('Please select a manufacturer or enter at least 2 characters for SKU search', 'error');
    }
}

/**
 * Look up manufacturers that have products matching the SKU pattern
 * @param {string} skuPattern - The SKU/part number pattern to search for
 */
async function lookupManufacturersFromSKU(skuPattern) {
    showStatus(`Searching for manufacturers with SKU matching "${skuPattern}"...`, 'loading');
    state.skuSearchMode = true;
    state.pendingSkuFilter = skuPattern;

    try {
        let manufacturers = [];

        if (state.currentDistributor === 'tdsynnex') {
            // TD SYNNEX: Call Supabase RPC get_manufacturers_by_sku
            const response = await fetch(`${SUPABASE_URL}/rest/v1/rpc/get_manufacturers_by_sku`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                },
                body: JSON.stringify({
                    sku_pattern: skuPattern
                })
            });

            if (!response.ok) {
                throw new Error(`RPC call failed: ${response.status}`);
            }

            const data = await response.json();
            manufacturers = data || [];

        } else {
            // Ingram: Search Supabase DB by SKU pattern across all manufacturers
            const encodedPattern = encodeURIComponent(`%${skuPattern}%`);
            const url = `${SUPABASE_URL}/rest/v1/zoho_ingram_products?select=manufacturer&or=(vendor_part_number.ilike.${encodedPattern},ingram_part_number.ilike.${encodedPattern})&manufacturer=not.is.null&limit=200`;
            const response = await fetch(url, {
                headers: {
                    'apikey': SUPABASE_ANON_KEY,
                    'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                }
            });
            const rows = await response.json();

            if (Array.isArray(rows)) {
                const vendorCounts = {};
                rows.forEach(row => {
                    const mfr = row.manufacturer || 'Unknown';
                    vendorCounts[mfr] = (vendorCounts[mfr] || 0) + 1;
                });

                manufacturers = Object.entries(vendorCounts).map(([name, count]) => ({
                    manufacturer_name: name,
                    product_count: count
                }));
            }
        }

        if (manufacturers.length === 0) {
            showStatus(`No manufacturers found with products matching "${skuPattern}"`, 'info');
            state.skuSearchMode = false;
            state.pendingSkuFilter = '';
            return;
        }

        if (manufacturers.length === 1) {
            // Single manufacturer - auto-select it
            const mfr = manufacturers[0];
            const mfrName = mfr.manufacturer_name || mfr;
            await handleSingleManufacturerAutoSelect(mfrName, skuPattern);
        } else {
            // Multiple manufacturers - populate dropdown for user selection
            populateManufacturerDropdownFromSKU(manufacturers);
        }

    } catch (error) {
        console.error('[SKU Search] Error:', error);
        showStatus('Error searching by SKU: ' + error.message, 'error');
        state.skuSearchMode = false;
        state.pendingSkuFilter = '';
    }
}

/**
 * Handle auto-selection when SKU search returns a single manufacturer
 * @param {string} manufacturer - The manufacturer name
 * @param {string} skuValue - The SKU value to use as filter
 */
async function handleSingleManufacturerAutoSelect(manufacturer, skuValue) {
    // Set state
    state.manufacturer = manufacturer;
    state.skuKeyword = skuValue;

    // Update manufacturer dropdown to show selected value
    const select = document.getElementById('manufacturerSelect');
    select.innerHTML = `<option value="${manufacturer}" selected>${manufacturer}</option>`;

    // Update manufacturer badge
    const mfrBadge = document.getElementById('selectedMfrBadge');
    if (mfrBadge) {
        mfrBadge.textContent = manufacturer;
    }

    // Hide OR divider
    const orDivider = document.getElementById('orDivider');
    if (orDivider) {
        orDivider.classList.add('hidden');
    }

    // Show optional filters
    document.getElementById('optionalFiltersRow').style.display = 'flex';

    // Keep the SKU value in the top search field and update placeholder
    const skuSearchField = document.getElementById('skuSearch');
    const skuSearchRow = document.getElementById('skuSearchRow');
    if (skuSearchField) {
        skuSearchField.value = skuValue;
        skuSearchField.placeholder = `Filter ${manufacturer} products by SKU...`;
    }
    if (skuSearchRow) {
        skuSearchRow.classList.add('filter-mode');
    }

    showStatus(`Auto-selected ${manufacturer}. Loading filters...`, 'loading');

    // Load filter options from the manufacturer
    await loadFilterOptions('category');
    if (state.currentDistributor === 'ingram') {
        await loadFilterOptions('skuType');
    }

    // Load products with the SKU filter
    await loadProducts(1);

    // Clear SKU search mode flags
    state.skuSearchMode = false;
    state.pendingSkuFilter = '';
}

/**
 * Populate the manufacturer dropdown with results from SKU search
 * @param {Array} manufacturers - Array of {manufacturer_name, product_count} objects
 */
function populateManufacturerDropdownFromSKU(manufacturers) {
    const select = document.getElementById('manufacturerSelect');

    // Clear and populate with SKU-matched manufacturers
    select.innerHTML = '<option value="">-- Select a manufacturer --</option>';

    manufacturers.forEach(mfr => {
        const name = mfr.manufacturer_name || mfr;
        const count = mfr.product_count || 0;
        const option = document.createElement('option');
        option.value = name;
        option.textContent = count > 0 ? `${name} (${count} matches)` : name;
        select.appendChild(option);
    });

    // Store options for reference
    state.skuManufacturerOptions = manufacturers;

    // Update count badge
    document.getElementById('mfrCount').textContent = `(${manufacturers.length})`;

    // Hide OR divider
    const orDivider = document.getElementById('orDivider');
    if (orDivider) {
        orDivider.classList.add('hidden');
    }

    showStatus(`Found ${manufacturers.length} manufacturers with matching SKUs. Please select one.`, 'info');
}

// =====================================================
// MANUFACTURER SELECTION
// =====================================================
async function onManufacturerSelect() {
    const select = document.getElementById('manufacturerSelect');
    state.manufacturer = select.value;

    // Hide OR divider when manufacturer is selected
    const orDivider = document.getElementById('orDivider');

    if (!state.manufacturer) {
        // Manufacturer cleared - reset state
        resetOptionalFilters();
        resetProducts();
        document.getElementById('optionalFiltersRow').style.display = 'none';
        document.getElementById('selectedMfrBadge').textContent = '';

        // Show OR divider when manufacturer is cleared
        if (orDivider) {
            orDivider.classList.remove('hidden');
        }

        // Reset SKU search mode and restore default placeholder
        state.skuSearchMode = false;
        state.pendingSkuFilter = '';
        state.skuManufacturerOptions = [];

        const skuSearchField = document.getElementById('skuSearch');
        const skuSearchRow = document.getElementById('skuSearchRow');
        if (skuSearchField) {
            skuSearchField.placeholder = 'Enter partial or full SKU (e.g. AB123, XYZ-456)...';
        }
        if (skuSearchRow) {
            skuSearchRow.classList.remove('filter-mode');
        }

        return;
    }

    // Hide OR divider when manufacturer is selected
    if (orDivider) {
        orDivider.classList.add('hidden');
    }

    // Check if we're in SKU search mode with a pending filter
    const hasPendingSkuFilter = state.skuSearchMode && state.pendingSkuFilter;

    if (!hasPendingSkuFilter) {
        // Normal manufacturer-first flow
        resetOptionalFilters();
        resetProducts();
    }

    document.getElementById('optionalFiltersRow').style.display = 'flex';

    // Update manufacturer badge
    const mfrBadge = document.getElementById('selectedMfrBadge');
    if (mfrBadge) {
        mfrBadge.textContent = state.manufacturer;
    }

    // Update SKU field placeholder to indicate filter mode
    const skuSearchField = document.getElementById('skuSearch');
    const skuSearchRow = document.getElementById('skuSearchRow');
    if (skuSearchField) {
        skuSearchField.placeholder = `Filter ${state.manufacturer} products by SKU...`;
    }
    if (skuSearchRow) {
        skuSearchRow.classList.add('filter-mode');
    }

    showStatus(`Manufacturer: ${state.manufacturer}. Loading filters...`, 'loading');

    await loadFilterOptions('category');
    if (state.currentDistributor === 'ingram') {
        await loadFilterOptions('skuType');
    }

    // If we have a pending SKU filter from SKU-first search, apply it
    if (hasPendingSkuFilter) {
        state.skuKeyword = state.pendingSkuFilter;

        // Keep the SKU value in the top search field
        if (skuSearchField) {
            skuSearchField.value = state.pendingSkuFilter;
        }

        // Load products with the SKU filter
        await loadProducts(1);

        // Clear SKU search mode flags
        state.skuSearchMode = false;
        state.pendingSkuFilter = '';
    } else {
        showStatus(`Manufacturer: ${state.manufacturer}. Use filters or click Load Products.`, 'success');
    }
}

// =====================================================
// FILTER LOADING
// =====================================================
async function loadFilterOptions(filterType) {
    const currentParams = `${state.manufacturer}|${state.brand}|${state.category}|${state.subcategory}|${state.cat3}|${state.skuType}`;

    if (state.loadingFilters[filterType]) return;
    if (state.filterParams[filterType] === currentParams) return;

    state.loadingFilters[filterType] = true;

    let selectEl, countEl;

    // Map filter type to DOM elements
    switch (filterType) {
        case 'brand':
            selectEl = document.getElementById('brandSelect');
            countEl = document.getElementById('brandCount');
            break;
        case 'category':
            selectEl = document.getElementById('categorySelect');
            countEl = document.getElementById('catCount');
            break;
        case 'subcategory':
            selectEl = document.getElementById('subcategorySelect');
            countEl = document.getElementById('subCatCount');
            break;
        case 'cat3':
            selectEl = document.getElementById('cat3Select');
            countEl = document.getElementById('cat3Count');
            break;
        case 'skuType':
            selectEl = document.getElementById('skuTypeSelect');
            countEl = document.getElementById('skuTypeCount');
            break;
        default:
            state.loadingFilters[filterType] = false;
            return;
    }

    const currentValue = selectEl.value;
    selectEl.innerHTML = '<option value="">Loading...</option>';

    try {
        let items = [];

        if (state.currentDistributor === 'tdsynnex') {
            // TD Synnex: Use dedicated edge function
            let categories = [];
            if (filterType === 'category') {
                categories = await loadTDSynnexCategories(state.manufacturer);
            } else if (filterType === 'subcategory' && state.category) {
                categories = await loadTDSynnexCategories(state.manufacturer, state.category);
            } else if (filterType === 'cat3' && state.category && state.subcategory) {
                categories = await loadTDSynnexCategories(state.manufacturer, state.category, state.subcategory);
            }
            items = categories.map(c => c.name);
        } else if (state.currentDistributor === 'adi') {
            // ADI Global: Only category filter supported (no subcategory, no cat3, no skuType)
            if (filterType === 'category') {
                const rpcBody = {
                    p_manufacturer: state.manufacturer || null
                };
                const response = await fetch(
                    `${SUPABASE_URL}/rest/v1/rpc/get_adi_categories`,
                    {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'apikey': SUPABASE_ANON_KEY,
                            'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                        },
                        body: JSON.stringify(rpcBody)
                    }
                );
                if (response.ok) {
                    const data = await response.json();
                    items = (data || []).map(r => r.category_1).filter(Boolean).sort();
                }
            } else {
                // ADI doesn't have subcategory, cat3, or skuType
                state.loadingFilters[filterType] = false;
                return;
            }
        } else {
            // Ingram: Query Supabase DB for categories/subcategories/media_type
            let rpcFilterType;
            switch (filterType) {
                case 'category':
                    rpcFilterType = 'category';
                    break;
                case 'subcategory':
                    rpcFilterType = 'subcategory';
                    break;
                case 'skuType':
                    rpcFilterType = 'media_type';
                    break;
                default:
                    // Ingram doesn't have cat3
                    state.loadingFilters[filterType] = false;
                    return;
            }

            const rpcBody = {
                p_manufacturer: state.manufacturer,
                p_filter_type: rpcFilterType,
                p_level1: state.category || null,
                p_level2: state.subcategory || null
            };

            const response = await fetch(
                `${SUPABASE_URL}/rest/v1/rpc/get_ingram_filter_values`,
                {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                    },
                    body: JSON.stringify(rpcBody)
                }
            );
            if (response.ok) {
                const data = await response.json();
                items = (data || []).map(r => r.value).filter(Boolean).sort();
            }
        }

        selectEl.innerHTML = '<option value="">-- Any --</option>';

        if (items.length > 0) {
            items.forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.textContent = item;
                selectEl.appendChild(option);
            });
            countEl.textContent = `(${items.length})`;

            if (currentValue && items.includes(currentValue)) {
                selectEl.value = currentValue;
            }
        } else {
            countEl.textContent = '(0)';
        }

        state.filterParams[filterType] = currentParams;

    } catch (error) {
        console.error(`Error loading ${filterType}:`, error);
        selectEl.innerHTML = '<option value="">-- Error --</option>';
    }

    state.loadingFilters[filterType] = false;
}

function formatSKUType(type) {
    switch (type) {
        case 'IM::physical':
        case 'IM::Physical':
        case 'Physical':
            return 'Physical';
        case 'IM::digital':
        case 'IM::Digital':
        case 'Digital':
            return 'Digital';
        case 'IM::subscription':
        case 'IM::Subscription':
        case 'Subscription':
            return 'Subscription';
        default:
            return type || '-';
    }
}

function formatProductClass(code) {
    const definitions = {
        'A': 'Stocked in all warehouses',
        'B': 'Stocked in limited warehouses',
        'C': 'Stocked in fewer warehouses',
        'D': 'Discontinued by Ingram',
        'E': 'Vendor phase-out',
        'F': 'Contract-specific product',
        'N': 'New SKU (pre-receipt)',
        'O': 'Discontinued - liquidation',
        'S': 'Special order / backorder',
        'X': 'Direct ship from vendor',
        'V': 'Discontinued by vendor'
    };

    if (!code) return '-';
    const upperCode = code.toUpperCase();
    const definition = definitions[upperCode];
    return definition ? `${upperCode} - ${definition}` : code;
}

async function onFilterChange(filterType) {
    const selectEl = document.getElementById(
        filterType === 'category' ? 'categorySelect' :
        filterType === 'subcategory' ? 'subcategorySelect' :
        filterType === 'cat3' ? 'cat3Select' :
        'skuTypeSelect'
    );

    state[filterType] = selectEl.value;

    // Reset dependent filters
    if (filterType !== 'category') state.filterParams.category = '';
    if (filterType !== 'subcategory') state.filterParams.subcategory = '';
    if (filterType !== 'cat3') state.filterParams.cat3 = '';
    if (filterType !== 'skuType') state.filterParams.skuType = '';

    // Clear downstream filters when parent changes
    if (filterType === 'category') {
        state.subcategory = '';
        state.cat3 = '';
        state.filterParams.subcategory = '';
        document.getElementById('subcategorySelect').innerHTML = '<option value="">-- Any --</option>';
        document.getElementById('subCatCount').textContent = '';
        if (state.currentDistributor === 'tdsynnex') {
            document.getElementById('cat3Select').innerHTML = '<option value="">-- Any --</option>';
            document.getElementById('cat3Count').textContent = '';
        }
    } else if (filterType === 'subcategory') {
        if (state.currentDistributor === 'tdsynnex') {
            state.cat3 = '';
            document.getElementById('cat3Select').innerHTML = '<option value="">-- Any --</option>';
            document.getElementById('cat3Count').textContent = '';
        }
    }

    resetProducts();

    // Load child/cross-filtered categories
    if (filterType === 'category') {
        // Reload subcategory (filtered by category for all distributors)
        await loadFilterOptions('subcategory');
        // Ingram: reload media types filtered by category
        if (state.currentDistributor === 'ingram') {
            state.skuType = '';
            const skuTypeSelect = document.getElementById('skuTypeSelect');
            if (skuTypeSelect) skuTypeSelect.value = '';
            await loadFilterOptions('skuType');
        }
    } else if (filterType === 'subcategory') {
        if (state.currentDistributor === 'tdsynnex') {
            await loadFilterOptions('cat3');
        } else if (state.currentDistributor === 'ingram') {
            // Ingram: reload media types filtered by category+subcategory
            state.skuType = '';
            const skuTypeSelect = document.getElementById('skuTypeSelect');
            if (skuTypeSelect) skuTypeSelect.value = '';
            await loadFilterOptions('skuType');
        }
    }
}

// =====================================================
// PRODUCTS LOADING
// =====================================================
async function loadProducts(page = 1) {
    if (!state.manufacturer) {
        showStatus('Please select a manufacturer first', 'error');
        return;
    }

    state.currentPage = page;
    const productsSection = document.getElementById('productsSection');
    productsSection.style.display = 'block';
    showStatus('Loading products with pricing...', 'loading');

    try {
        let products = [];
        let pagination = null;

        if (state.currentDistributor === 'tdsynnex') {
            // TD Synnex: Use dedicated edge function
            const offset = (page - 1) * PAGE_SIZE;
            const result = await searchTDSynnexProducts(state.manufacturer, {
                search: state.skuKeyword || '',
                cat1: state.category || '',
                cat2: state.subcategory || '',
                cat3: state.cat3 || '',
                limit: PAGE_SIZE,
                offset: offset
            });
            products = result.products;
            pagination = result.pagination;
        } else if (state.currentDistributor === 'adi') {
            // ADI Global: Query Supabase DB
            const offset = (page - 1) * PAGE_SIZE;
            const rpcBody = {
                p_manufacturer: state.manufacturer,
                p_search: (state.skuKeyword && state.skuKeyword.length >= 2) ? state.skuKeyword : null,
                p_category_1: state.category || null,
                p_limit: PAGE_SIZE,
                p_offset: offset
            };

            // Fetch products and count in parallel
            const [productsResponse, countResponse] = await Promise.all([
                fetch(`${SUPABASE_URL}/rest/v1/rpc/search_adi_products`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                    },
                    body: JSON.stringify(rpcBody)
                }),
                fetch(`${SUPABASE_URL}/rest/v1/rpc/search_adi_products_count`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                    },
                    body: JSON.stringify({
                        p_manufacturer: rpcBody.p_manufacturer,
                        p_search: rpcBody.p_search,
                        p_category_1: rpcBody.p_category_1
                    })
                })
            ]);

            const rows = await productsResponse.json();
            const totalRecords = await countResponse.json();

            products = (rows || []).map(row => mapADIGlobalProduct(row));

            const totalPages = Math.ceil((totalRecords || 0) / PAGE_SIZE);
            pagination = {
                page: page,
                pageSize: PAGE_SIZE,
                totalPages: totalPages,
                totalRecords: totalRecords || 0
            };
        } else {
            // Ingram: Query Supabase DB
            const offset = (page - 1) * PAGE_SIZE;
            const rpcBody = {
                p_manufacturer: state.manufacturer,
                p_search: (state.skuKeyword && state.skuKeyword.length >= 2) ? state.skuKeyword : null,
                p_level1: state.category || null,
                p_level2: state.subcategory || null,
                p_media_type: state.skuType || null,
                p_limit: PAGE_SIZE,
                p_offset: offset
            };

            // Fetch products and count in parallel
            const [productsResponse, countResponse] = await Promise.all([
                fetch(`${SUPABASE_URL}/rest/v1/rpc/search_ingram_products`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                    },
                    body: JSON.stringify(rpcBody)
                }),
                fetch(`${SUPABASE_URL}/rest/v1/rpc/search_ingram_products_count`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'apikey': SUPABASE_ANON_KEY,
                        'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
                    },
                    body: JSON.stringify({
                        p_manufacturer: rpcBody.p_manufacturer,
                        p_search: rpcBody.p_search,
                        p_level1: rpcBody.p_level1,
                        p_level2: rpcBody.p_level2,
                        p_media_type: rpcBody.p_media_type
                    })
                })
            ]);

            const rows = await productsResponse.json();
            const totalRecords = await countResponse.json();

            // Map DB columns to widget product format
            products = (rows || []).map(row => ({
                ingramPartNumber: row.ingram_part_number || '',
                vendorPartNumber: row.vendor_part_number || '',
                vendorName: row.manufacturer || '',
                description: row.description_line_1 || '',
                extraDescription: [row.description_line_1, row.description_line_2].filter(Boolean).join(' '),
                category: row.level_1_name || '',
                subCategory: row.level_2_name || '',
                productType: row.media_type || '',
                type: row.media_type || '',
                replacementSku: row.substitute_part_number || '',
                upcCode: row.upc_code || '',
                availabilityFlag: row.availability_flag || '',
                status: row.status || '',
                cpuCode: row.cpu_code || '',
                // Construct pricingData from DB prices (weekly refresh — not live)
                // _dbSource flag prevents caching in state.pricingData so showProductDetails
                // always fetches live pricing/availability from the Ingram API
                pricingData: {
                    _dbSource: true,
                    pricing: {
                        retailPrice: row.retail_price ? parseFloat(row.retail_price) : null,
                        customerPrice: row.customer_price ? parseFloat(row.customer_price) : null
                    }
                },
                resellerPrice: row.customer_price ? parseFloat(row.customer_price) : null,
                _source: 'ingram'
            }));

            const totalPages = Math.ceil((totalRecords || 0) / PAGE_SIZE);
            pagination = {
                page: page,
                pageSize: PAGE_SIZE,
                totalPages: totalPages,
                totalRecords: totalRecords || 0
            };
        }

        if (products.length > 0) {
            // Store total records for pagination display
            state.totalRecords = pagination?.totalRecords || products.length;
            state.totalPages = pagination?.totalPages || 1;

            displayProductsWithPricing(products, pagination);
            showStatus('', '');
            document.getElementById('productsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
        } else {
            document.getElementById('productsBody').innerHTML =
                '<tr><td colspan="5" class="no-results">No products found</td></tr>';
            document.getElementById('pagination').innerHTML = '';
            document.getElementById('productCount').textContent = '0 products';
            showStatus('No products found with current filters', 'info');
        }
    } catch (error) {
        showStatus('Error loading products: ' + error.message, 'error');
    }
}

function displayProductsWithPricing(products, pagination) {
    const tbody = document.getElementById('productsBody');
    tbody.innerHTML = '';

    const sortedProducts = [...products].sort((a, b) => {
        const partA = (a.vendorPartNumber || '').toLowerCase();
        const partB = (b.vendorPartNumber || '').toLowerCase();
        return partA.localeCompare(partB, undefined, { numeric: true, sensitivity: 'base' });
    });

    state.currentProducts = sortedProducts;
    state.pricingData = {};

    // Lazy verification — fire-and-forget, non-blocking
    if (state.currentDistributor === 'ingram') {
        verifyIngramManufacturers(sortedProducts);
    }

    sortedProducts.forEach((product, index) => {
        const partNumber = product.ingramPartNumber || product.vendorPartNumber;
        const isSelected = state.selectedProducts.has(partNumber);
        const isQueued = state.queuedProducts.some(p =>
            (p.ingramPartNumber || p.vendorPartNumber) === partNumber
        );

        const pricingData = product.pricingData;
        const msrp = pricingData?.pricing?.retailPrice;
        const msrpDisplay = msrp
            ? `<span class="price-available">$${msrp.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>`
            : '<span class="price-unavailable">-</span>';

        if (pricingData && product.ingramPartNumber && !pricingData._dbSource) {
            state.pricingData[product.ingramPartNumber] = pricingData;
        }

        const tr = document.createElement('tr');
        tr.className = isSelected ? 'selected' : '';
        if (isQueued) tr.classList.add('queued');
        tr.id = `product-row-${index}`;

        const fullDescription = product.description || '-';
        // Simplified table: Checkbox, Part Number, Description (with hover tooltip), MSRP, Info
        tr.innerHTML = `
            <td class="col-checkbox">
                <input type="checkbox"
                       onchange="toggleProduct('${partNumber}', this.checked)"
                       ${isSelected ? 'checked' : ''}
                       ${isQueued ? 'disabled title="Already in queue"' : ''}>
            </td>
            <td class="col-part"><strong>${product.vendorPartNumber || '-'}</strong></td>
            <td class="col-desc desc-cell" title="${fullDescription.replace(/"/g, '&quot;')}">${fullDescription}</td>
            <td class="col-price">${msrpDisplay}</td>
            <td class="col-action">
                <button class="info-btn" onclick="showProductDetails(${index})" title="View details">i</button>
            </td>
        `;
        tbody.appendChild(tr);

        tr.dataset.product = JSON.stringify(product);
    });

    // Use stored total records for accurate count across all pages
    document.getElementById('productCount').textContent =
        `${state.totalRecords.toLocaleString()} products`;

    renderPagination(pagination);
    updateSelectedCount();
}

function renderPagination(pagination) {
    const paginationDiv = document.getElementById('pagination');

    if (!pagination || pagination.totalPages <= 1) {
        paginationDiv.innerHTML = '';
        return;
    }

    paginationDiv.innerHTML = `
        <button onclick="loadProducts(${pagination.page - 1})"
                ${pagination.page === 1 ? 'disabled' : ''} class="btn-secondary btn-small">
            Prev
        </button>
        <span>Page ${pagination.page} of ${pagination.totalPages} (${state.totalRecords.toLocaleString()} total)</span>
        <button onclick="loadProducts(${pagination.page + 1})"
                ${pagination.page >= pagination.totalPages ? 'disabled' : ''} class="btn-secondary btn-small">
            Next
        </button>
    `;
}

// =====================================================
// PRODUCT SELECTION
// =====================================================
function toggleProduct(partNumber, isChecked) {
    const rows = document.querySelectorAll('#productsBody tr');

    rows.forEach(row => {
        const productData = row.dataset.product;
        if (productData) {
            const product = JSON.parse(productData);
            const pn = product.ingramPartNumber || product.vendorPartNumber;

            if (pn === partNumber) {
                if (isChecked) {
                    state.selectedProducts.set(partNumber, product);
                    row.classList.add('selected');
                } else {
                    state.selectedProducts.delete(partNumber);
                    row.classList.remove('selected');
                }
            }
        }
    });

    updateSelectedCount();
}

function toggleSelectAll() {
    const selectAllChecked = document.getElementById('selectAll').checked;
    const checkboxes = document.querySelectorAll('#productsBody input[type="checkbox"]:not(:disabled)');

    checkboxes.forEach(cb => {
        cb.checked = selectAllChecked;
        const row = cb.closest('tr');
        const productData = row.dataset.product;

        if (productData) {
            const product = JSON.parse(productData);
            const partNumber = product.ingramPartNumber || product.vendorPartNumber;

            if (selectAllChecked) {
                state.selectedProducts.set(partNumber, product);
                row.classList.add('selected');
            } else {
                state.selectedProducts.delete(partNumber);
                row.classList.remove('selected');
            }
        }
    });

    updateSelectedCount();
}

function updateFooterStats() {
    var el = document.getElementById('footerStats');
    if (!el) return;

    if (state.searchMode === 'bulk') {
        var parsed = bulkState.parsedSkus ? bulkState.parsedSkus.length : 0;
        var found = bulkState.products ? bulkState.products.length : 0;
        var notFound = bulkState.unmatchedMpns ? bulkState.unmatchedMpns.length : 0;
        var selected = bulkState.selectedProductIndices ? bulkState.selectedProductIndices.size : 0;
        var hasResults = found > 0 || notFound > 0;

        if (!parsed && !hasResults) {
            el.innerHTML = '';
        } else if (parsed > 0 && !hasResults) {
            el.innerHTML = '<strong>' + parsed + '</strong> SKUs parsed';
        } else if (hasResults && selected > 0) {
            el.innerHTML = '<strong>' + found + '</strong> of ' + parsed + ' found · <strong>' + notFound + '</strong> not found · <strong>' + selected + '</strong> selected';
        } else if (hasResults) {
            el.innerHTML = '<strong>' + found + '</strong> of ' + parsed + ' found · <strong>' + notFound + '</strong> not found';
        }
    } else {
        var count = state.selectedProducts ? state.selectedProducts.size : 0;
        el.innerHTML = '<strong>' + count + '</strong> selected from search';
    }
}

function updateSelectedCount() {
    var count = state.selectedProducts.size;
    var addToQueueBtn = document.getElementById('addToQueueBtn');
    if (addToQueueBtn) {
        addToQueueBtn.disabled = count === 0;
    }
    updateFooterStats();
}

function updateFooterDateTime() {
    var el = document.getElementById('footerDateTime');
    if (!el) return;
    var now = new Date();
    var options = {
        timeZone: 'America/Denver',
        month: '2-digit',
        day: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
    };
    var formatted = now.toLocaleString('en-US', options).replace(',', ' -');
    el.textContent = formatted;
}

// Start footer clock — update every 30 seconds
updateFooterDateTime();
setInterval(updateFooterDateTime, 30000);

// =====================================================
// QUEUE MANAGEMENT
// =====================================================
function addSelectedToQueue() {
    const selectedArray = Array.from(state.selectedProducts.values());

    if (selectedArray.length === 0) {
        showStatus('No products selected', 'error');
        return;
    }

    let addedCount = 0;
    selectedArray.forEach(product => {
        const partNumber = product.ingramPartNumber || product.vendorPartNumber;
        const alreadyQueued = state.queuedProducts.some(p =>
            (p.ingramPartNumber || p.vendorPartNumber) === partNumber
        );

        if (!alreadyQueued) {
            // Enrich product with pricing data if available
            const pricingData = product.pricingData || state.pricingData?.[product.ingramPartNumber];
            const enrichedProduct = { ...product, pricingData, customerDiscount: 0 };
            state.queuedProducts.push(enrichedProduct);
            addedCount++;
        }
    });

    // Clear current selection
    state.selectedProducts.clear();
    updateSelectedCount();

    // Uncheck all checkboxes
    document.querySelectorAll('#productsBody input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
        cb.closest('tr').classList.remove('selected');
    });
    document.getElementById('selectAll').checked = false;

    // Refresh product display to show queued items as disabled
    if (state.currentProducts.length > 0) {
        displayProductsWithPricing(state.currentProducts, {
            totalRecords: state.totalRecords,
            page: state.currentPage,
            totalPages: state.totalPages
        });
    }

    updateQueueUI();
    document.querySelector('.queue-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });

    if (addedCount > 0) {
        showStatus(`Added ${addedCount} product(s) to queue`, 'success');
    } else {
        showStatus('Products already in queue', 'info');
    }
}

function updateQueueItemQty(mpn, value) {
    const product = getActiveQueue().find(p =>
        (p.ingramPartNumber || p.vendorPartNumber) === mpn
    );
    if (!product) return;
    let qty = parseInt(value, 10);
    if (isNaN(qty) || qty < 1) qty = 1;
    if (qty > 9999) qty = 9999;
    product.qty = qty;
    try { if (typeof updateQueueTotals === 'function') updateQueueTotals(); } catch (e) { /* Task 4 */ }
}

function removeFromQueue(partNumber) {
    setActiveQueue(getActiveQueue().filter(p =>
        (p.ingramPartNumber || p.vendorPartNumber) !== partNumber
    ));
    updateQueueUI();

    // Re-enable checkbox in products table if visible
    document.querySelectorAll('#productsBody input[type="checkbox"][disabled]').forEach(cb => {
        const row = cb.closest('tr');
        const productData = row.dataset.product;
        if (productData) {
            const product = JSON.parse(productData);
            const pn = product.ingramPartNumber || product.vendorPartNumber;
            if (pn === partNumber) {
                cb.disabled = false;
                cb.title = '';
                row.classList.remove('queued');
            }
        }
    });
}

function clearQueue() {
    setActiveQueue([]);
    updateQueueUI();

    // Re-enable all disabled checkboxes
    document.querySelectorAll('#productsBody input[type="checkbox"][disabled]').forEach(cb => {
        cb.disabled = false;
        cb.title = '';
        cb.closest('tr').classList.remove('queued');
    });

    showStatus('Queue cleared', 'info');
}

function getActiveQueue() {
    return state.searchMode === 'bulk' ? bulkState.queuedProducts : state.queuedProducts;
}

function setActiveQueue(newArray) {
    if (state.searchMode === 'bulk') {
        bulkState.queuedProducts = newArray;
    } else {
        state.queuedProducts = newArray;
    }
}

function getActivePricingMode() {
    return state.searchMode === 'bulk' ? bulkState.pricingMode : state.pricingMode;
}

function updateQueueTotals() {
    const totalQtyEl = document.getElementById('totalQty');
    const totalPriceEl = document.getElementById('totalPrice');
    const totalPriceLabelEl = document.getElementById('totalPriceLabel');
    if (!totalQtyEl || !totalPriceEl) return;

    const totalQty = getActiveQueue().reduce((sum, item) => sum + (item.qty || 1), 0);

    const totalPrice = getActiveQueue().reduce((sum, item) => {
        const qty = item.qty || 1;
        const msrp = item.pricingData?.pricing?.retailPrice || item.retailPrice || item.msrp || 0;
        const resellerPrice = item.resellerPrice || item.pricingData?.pricing?.customerPrice || null;
        const price = (getActivePricingMode() === 'reseller' && resellerPrice !== null) ? resellerPrice : msrp;
        return sum + (qty * price);
    }, 0);

    totalQtyEl.textContent = totalQty.toLocaleString();
    totalPriceEl.textContent = '$' + totalPrice.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    totalPriceLabelEl.textContent = getActivePricingMode() === 'reseller' ? 'Total Reseller' : 'Total MSRP';
}

function updateQueueUI() {
    const queueCount = getActiveQueue().length;

    document.getElementById('queueCount').textContent = queueCount;

    const queueEmpty = document.getElementById('queueEmpty');
    const queueList = document.getElementById('queueList');
    const queueFooter = document.getElementById('queueFooter');
    const clearQueueBtn = document.getElementById('clearQueueBtn');
    const queueOptions = document.getElementById('queueOptions');
    const queueTotals = document.getElementById('queueTotals');

    if (queueCount === 0) {
        queueEmpty.style.display = 'flex';
        queueList.style.display = 'none';
        queueFooter.style.display = 'none';
        clearQueueBtn.style.display = 'none';
        if (queueOptions) queueOptions.style.display = 'none';
        if (queueTotals) queueTotals.style.display = 'none';
    } else {
        queueEmpty.style.display = 'none';
        queueList.style.display = 'block';
        queueFooter.style.display = 'block';
        clearQueueBtn.style.display = 'block';
        if (queueOptions) queueOptions.style.display = 'block';
        if (queueTotals) queueTotals.style.display = 'flex';

        renderQueueItems();
        updateQueueTotals();
    }
}

function renderQueueItems() {
    const queueItems = document.getElementById('queueItems');
    queueItems.innerHTML = '';

    if (state.groupByManufacturer) {
        // Group products by manufacturer
        const groups = {};
        getActiveQueue().forEach(product => {
            const mfr = product.vendorName || product.manufacturer || 'Unknown';
            if (!groups[mfr]) {
                groups[mfr] = [];
            }
            groups[mfr].push(product);
        });

        // Render grouped (preserve insertion order, not alphabetical)
        Object.keys(groups).forEach(mfr => {
            // Add manufacturer header (draggable)
            const header = document.createElement('div');
            header.className = 'queue-mfr-group';
            header.innerHTML = '<span class="queue-mfr-group-left">' + escapeHtml(mfr) + '</span>' +
                    '<span class="queue-mfr-group-discount">' +
                        '<input type="number" class="mfr-discount-input" value="0" min="0" max="100" step="1" ' +
                            'placeholder="0" data-mfr="' + mfr.replace(/"/g, '&quot;') + '" ' +
                            'onclick="event.stopPropagation(); this.select()" ' +
                            'onkeydown="if(event.key===\'Enter\'){event.stopPropagation(); applyMfrDiscount(\'' + mfr.replace(/'/g, "\\'") + '\'); this.blur();}" ' +
                            'onblur="applyMfrDiscountIfNonZero(\'' + mfr.replace(/'/g, "\\'") + '\')">' +
                        '<span class="mfr-discount-pct">%</span>' +
                        '<button class="mfr-discount-apply" data-tooltip="Apply to group" onclick="event.stopPropagation(); applyMfrDiscount(\'' + mfr.replace(/'/g, "\\'") + '\')">' +
                            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 13l4 4L19 7"/></svg>' +
                        '</button>' +
                    '</span>';
            header.style.display = 'flex';
            header.style.justifyContent = 'space-between';
            header.style.alignItems = 'center';
            header.style.textAlign = 'left';
            header.draggable = true;
            header.dataset.manufacturer = mfr;
            queueItems.appendChild(header);
                queueItems.insertAdjacentHTML('beforeend', createQueueColumnLabelsHTML());

            // Add items for this manufacturer
            groups[mfr].forEach((product, index) => {
                queueItems.appendChild(createQueueItemElement(product, index));
            });
        });
    } else {
        // Render flat list
        queueItems.insertAdjacentHTML('beforeend', createQueueColumnLabelsHTML());
        getActiveQueue().forEach((product, index) => {
            queueItems.appendChild(createQueueItemElement(product, index));
        });
    }
}

function createQueueItemElement(product, index) {
    const partNumber = product.ingramPartNumber || product.vendorPartNumber;
    const msrp = product.pricingData?.pricing?.retailPrice || product.retailPrice || product.msrp;
    const resellerPrice = product.resellerPrice || product.pricingData?.pricing?.customerPrice || null;
    const displayPrice = (getActivePricingMode() === 'reseller' && resellerPrice !== null) ? resellerPrice : msrp;
    const msrpDisplay = displayPrice
        ? `$${displayPrice.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
        : '-';

    const li = document.createElement('li');
    li.className = 'queue-item';
    li.draggable = true;
    li.dataset.partNumber = partNumber;
    li.dataset.index = index;

    // Tooltip data
    const description = product.description || product.descriptionLine || '';
    const tooltipMsrp = msrp
        ? `$${Number(msrp).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
        : '-';
    const tooltipReseller = resellerPrice
        ? `$${Number(resellerPrice).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
        : 'N/A';

    // Minimal: drag handle, part number, price, remove button
    li.innerHTML = `
        <div class="queue-item-tooltip">
            <div class="tooltip-desc">${description}</div>
            <div class="tooltip-prices">
                <div class="tooltip-price-item">
                    <span class="tooltip-price-label">MSRP</span>
                    <span class="tooltip-price-value">${tooltipMsrp}</span>
                </div>
                <div class="tooltip-price-item">
                    <span class="tooltip-price-label">Reseller</span>
                    <span class="tooltip-price-value">${tooltipReseller}</span>
                </div>
            </div>
        </div>
        <div class="queue-item-drag">
            <svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="9" cy="5" r="1"/><circle cx="9" cy="12" r="1"/><circle cx="9" cy="19" r="1"/>
                <circle cx="15" cy="5" r="1"/><circle cx="15" cy="12" r="1"/><circle cx="15" cy="19" r="1"/>
            </svg>
        </div>
        <input type="number" class="queue-item-qty" value="${product.qty || 1}" min="1" max="9999"
               onchange="updateQueueItemQty('${partNumber.replace(/'/g, "\\'")}', this.value)"
               onclick="event.stopPropagation()">
        <div class="queue-item-info">
            <div class="queue-item-part">${product.vendorPartNumber || '-'}</div>
        </div>
        ${state.searchMode === 'bulk'
            ? `<span class="queue-item-msrp-indicator">${product._msrpAdjusted
                ? (getActivePricingMode() === 'msrp'
                    ? `<span class="bulk-msrp-indicator bulk-msrp-${product._msrpDirection}">${product._msrpDirection === 'down' ? '▼' : '▲'}</span>`
                    : `<span class="bulk-msrp-dot bulk-msrp-${product._msrpDirection}">●</span>`)
                : ''}</span>`
            : ''}
        <div class="queue-item-price">${msrpDisplay}</div>
        <div class="queue-item-discount-wrap">
            <input type="number" class="queue-item-discount${product.customerDiscount > 0 ? ' has-value' : ''}" value="${product.customerDiscount || 0}" min="0" max="100" step="1" data-mpn="${partNumber}" onclick="event.stopPropagation(); this.select()" oninput="updateQueueItemDiscount('${partNumber.replace(/'/g, "\\'")}', this)" onblur="finalizeQueueItemDiscount('${partNumber.replace(/'/g, "\\'")}', this)">
            <span class="queue-item-discount-pct">%</span>
        </div>
        <button class="queue-item-remove" onclick="removeFromQueue('${partNumber}')" title="Remove">
            <svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M18 6 6 18M6 6l12 12"/>
            </svg>
        </button>
    `;

    return li;
}

// =====================================================
// MANUFACTURER NORMALIZATION
// =====================================================
async function normalizeManufacturer(name, distributor) {
    if (!name) return name;
    try {
        console.log(`[MfrNorm] Calling RPC for: "${name}" (${distributor})`);
        const response = await fetch(`${SUPABASE_URL}/rest/v1/rpc/normalize_manufacturer_name`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify({
                input_name: name,
                source_distributor: distributor
            })
        });
        console.log(`[MfrNorm] Response status: ${response.status}`);
        const result = await response.json();
        console.log(`[MfrNorm] Result: ${JSON.stringify(result)}`);
        if (response.ok && result) {
            console.log(`[MfrNorm] SUCCESS: "${name}" -> "${result}"`);
            return result;
        }
        console.log(`[MfrNorm] No result, returning original: "${name}"`);
        return name;
    } catch (error) {
        console.error('[MfrNorm] ERROR:', error);
        return name;
    }
}

// =====================================================
// MANUFACTURER RESOLUTION - BATCH RPC FUNCTIONS
// =====================================================

/**
 * Check manufacturer mappings in batch via Supabase RPC
 * @param {Array} manufacturers - Array of {name: string, distributor: string}
 * @returns {Promise<Array>} - Array of {name, distributor, found, canonical_name}
 */
async function checkManufacturerMappingsBatch(manufacturers) {
    if (!manufacturers || manufacturers.length === 0) return [];

    try {
        console.log('[MfrResolution] Checking mappings for:', manufacturers);
        const response = await fetch(`${SUPABASE_URL}/rest/v1/rpc/check_manufacturer_mappings_batch`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify({
                p_mappings: manufacturers
            })
        });

        if (!response.ok) {
            throw new Error(`RPC call failed: ${response.status}`);
        }

        const results = await response.json();
        console.log('[MfrResolution] Mapping check results:', results);
        return results || [];
    } catch (error) {
        console.error('[MfrResolution] Error checking mappings:', error);
        throw error;
    }
}

/**
 * Save manufacturer mappings in batch via Supabase RPC
 * @param {Array} mappings - Array of {distributor_name, canonical_name, source}
 * @returns {Promise<Object>} - {success: boolean, saved_count: number}
 */
async function saveManufacturerMappingsBatch(mappings) {
    if (!mappings || mappings.length === 0) return { success: true, saved_count: 0 };

    try {
        console.log('[MfrResolution] Saving mappings:', mappings);
        const response = await fetch(`${SUPABASE_URL}/rest/v1/rpc/save_manufacturer_mappings_batch`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            },
            body: JSON.stringify({
                p_mappings: mappings
            })
        });

        if (!response.ok) {
            throw new Error(`RPC call failed: ${response.status}`);
        }

        const result = await response.json();
        console.log('[MfrResolution] Save result:', result);
        return result || { success: false, saved_count: 0 };
    } catch (error) {
        console.error('[MfrResolution] Error saving mappings:', error);
        throw error;
    }
}

// =====================================================
// MANUFACTURER RESOLUTION - PANEL FUNCTIONS
// =====================================================

/**
 * Show the manufacturer resolution panel with unresolved manufacturers
 * @param {Array} unresolvedList - Array of {distributorName, distributor}
 * @returns {Promise} - Resolves with normalized map or rejects on cancel
 */
async function showMfrResolutionPanel(unresolvedList) {
    // If no prefetched manufacturers, try to fetch from Zoho directly
    if (state.prefetchedManufacturers.length === 0 && typeof ZOHO !== 'undefined') {
        console.log('[MfrResolution] No prefetched manufacturers, attempting direct Zoho fetch...');
        try {
            const response = await ZOHO.CRM.API.getAllRecords({
                Entity: "Manufacturers",
                sort_by: "Name",
                sort_order: "asc",
                per_page: 200
            });

            if (response && response.data && Array.isArray(response.data)) {
                state.prefetchedManufacturers = response.data
                    .map(record => record.Name || '')
                    .filter(name => name.length > 0)
                    .sort();
                console.log(`[MfrResolution] Fetched ${state.prefetchedManufacturers.length} manufacturers directly from Zoho`);
            }
        } catch (error) {
            console.warn('[MfrResolution] Failed to fetch manufacturers from Zoho:', error);
        }
    }

    return new Promise((resolve, reject) => {
        state.unresolvedManufacturers = unresolvedList;
        state.mfrResolutions = new Map();
        state.mfrResolutionPromise = { resolve, reject };

        renderMfrResolutionTable();
        updateMfrResolutionStatus();

        // Initialize tooltip positioning after table is rendered
        initMfrResolutionTooltips();

        const panel = document.getElementById('mfrResolutionPanel');
        if (panel) {
            // In bulk mode, move panel into the bulk search area
            if (state.searchMode === 'bulk') {
                const anchor = document.getElementById('bulkMfrResolutionAnchor');
                if (anchor) anchor.appendChild(panel);
            }
            panel.style.display = 'block';
            panel.classList.remove('collapsed');
            panel.classList.remove('all-resolved');
            panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }

        const collapseBtn = document.getElementById('mfrCollapseBtn');
        if (collapseBtn) {
            collapseBtn.classList.remove('collapsed');
        }
    });
}

/**
 * Hide the manufacturer resolution panel
 */
function hideMfrResolutionPanel() {
    const panel = document.getElementById('mfrResolutionPanel');
    if (panel) {
        panel.style.display = 'none';
        // If panel was moved to bulk anchor, move it back to original location
        const bulkAnchor = document.getElementById('bulkMfrResolutionAnchor');
        if (bulkAnchor && bulkAnchor.contains(panel)) {
            // Move back to original parent (after the details-panel placeholder in single-search-panel)
            const singlePanel = document.querySelector('.single-search-panel');
            if (singlePanel) singlePanel.appendChild(panel);
        }
    }
    state.unresolvedManufacturers = [];
    state.mfrResolutions = new Map();
    state.mfrResolutionPromise = null;
}

/**
 * Toggle panel collapse state
 */
function toggleMfrPanelCollapse() {
    const panel = document.getElementById('mfrResolutionPanel');
    const collapseBtn = document.getElementById('mfrCollapseBtn');

    if (panel && collapseBtn) {
        panel.classList.toggle('collapsed');
        collapseBtn.classList.toggle('collapsed');
    }
}

/**
 * Render the resolution table rows with distributor group separators
 */
function renderMfrResolutionTable() {
    const tbody = document.getElementById('mfrResolutionTableBody');
    if (!tbody) return;

    // Debug: Log prefetchedManufacturers
    console.log('[MfrResolution] Rendering table with', state.prefetchedManufacturers.length, 'Zoho manufacturers');

    const zohoOptions = state.prefetchedManufacturers
        .map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`)
        .join('');

    // Group manufacturers by distributor for group separators
    let html = '';
    let currentDistributor = '';

    state.unresolvedManufacturers.forEach((mfr, index) => {
        // Debug: Log manufacturer data to verify correct fields
        console.log(`[MfrResolution] Row ${index}: name="${mfr.distributorName}", distributor="${mfr.distributor}"`);

        // Determine distributor label and class based on source
        // mfr.distributor should be 'ingram' or 'tdsynnex'
        const distributorLabel = mfr.distributor === 'ingram' ? 'INGRAM MICRO' :
                                  mfr.distributor === 'tdsynnex' ? 'TD SYNNEX' :
                                  'UNKNOWN';
        const distributorClass = mfr.distributor === 'ingram' ? 'ingram' :
                                  mfr.distributor === 'tdsynnex' ? 'tdsynnex' :
                                  'unknown';

        // Add distributor group separator when distributor changes
        if (mfr.distributor !== currentDistributor) {
            currentDistributor = mfr.distributor;
            const count = state.unresolvedManufacturers.filter(m => m.distributor === currentDistributor).length;
            html += `
                <tr class="mfr-group-separator">
                    <td colspan="4">
                        <div class="mfr-group-label">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4"/>
                            </svg>
                            ${distributorLabel}
                            <span class="mfr-group-count">${count}</span>
                        </div>
                    </td>
                </tr>
            `;
        }

        // Add manufacturer row
        html += `
            <tr id="mfr-row-${index}" class="mfr-row">
                <td class="col-source">
                    <div class="mfr-distributor-cell">
                        <span class="mfr-name-cell">${escapeHtml(mfr.distributorName)}</span>
                    </div>
                </td>
                <td>
                    <div class="mfr-select-wrapper">
                        <select
                            class="mfr-select"
                            id="mfr-select-${index}"
                            data-index="${index}"
                            onchange="handleMfrSelectChange(${index})"
                        >
                            <option value="">-- Select manufacturer --</option>
                            ${zohoOptions}
                        </select>
                        <span class="mfr-select-arrow">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M6 9l6 6 6-6"/>
                            </svg>
                        </span>
                    </div>
                </td>
                <td class="td-or"></td>
                <td>
                    <div class="mfr-input-wrapper">
                        <input
                            type="text"
                            class="mfr-input"
                            id="mfr-input-${index}"
                            data-index="${index}"
                            placeholder="Enter new name..."
                            oninput="handleMfrInputChange(${index})"
                        />
                        <span id="mfr-status-${index}" class="mfr-row-status">
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                <path d="M5 12l5 5L20 7"/>
                            </svg>
                        </span>
                    </div>
                </td>
            </tr>
        `;
    });

    tbody.innerHTML = html;
}

/**
 * Handle select dropdown change
 */
function handleMfrSelectChange(index) {
    const select = document.getElementById(`mfr-select-${index}`);
    const input = document.getElementById(`mfr-input-${index}`);
    const row = document.getElementById(`mfr-row-${index}`);
    const status = document.getElementById(`mfr-status-${index}`);

    if (select.value) {
        // Disable input, set resolution
        input.value = '';
        input.disabled = true;
        state.mfrResolutions.set(index, { type: 'zoho', value: select.value });
        row.classList.add('row-valid');
        status.classList.add('show', 'valid');
    } else {
        // Re-enable input
        input.disabled = false;
        state.mfrResolutions.delete(index);
        row.classList.remove('row-valid');
        status.classList.remove('show', 'valid');
    }

    updateMfrResolutionStatus();
}

/**
 * Handle text input change
 */
function handleMfrInputChange(index) {
    const select = document.getElementById(`mfr-select-${index}`);
    const input = document.getElementById(`mfr-input-${index}`);
    const row = document.getElementById(`mfr-row-${index}`);
    const status = document.getElementById(`mfr-status-${index}`);

    // Auto-convert to Title Case while preserving cursor position
    const cursorPos = input.selectionStart;
    const originalLength = input.value.length;
    input.value = toTitleCase(input.value);
    const newLength = input.value.length;
    // Restore cursor position (adjust if length changed, though it shouldn't for title case)
    input.setSelectionRange(cursorPos + (newLength - originalLength), cursorPos + (newLength - originalLength));

    if (input.value.trim()) {
        // Disable select, set resolution
        select.value = '';
        select.disabled = true;
        state.mfrResolutions.set(index, { type: 'new', value: input.value.trim() });
        row.classList.add('row-valid');
        status.classList.add('show', 'valid');
    } else {
        // Re-enable select
        select.disabled = false;
        state.mfrResolutions.delete(index);
        row.classList.remove('row-valid');
        status.classList.remove('show', 'valid');
    }

    updateMfrResolutionStatus();
}

/**
 * Update the resolution status display
 */
function updateMfrResolutionStatus() {
    const total = state.unresolvedManufacturers.length;
    const resolved = state.mfrResolutions.size;
    const allResolved = resolved === total && total > 0;

    // Update count badge
    const countEl = document.getElementById('mfrUnresolvedCount');
    if (countEl) {
        countEl.textContent = total - resolved || total;
    }

    // Update status text
    const statusText = document.getElementById('mfrStatusText');
    if (statusText) {
        statusText.textContent = `${resolved} of ${total} resolved`;
    }

    // Update status dot
    const statusDot = document.getElementById('mfrStatusDot');
    if (statusDot) {
        statusDot.classList.toggle('complete', allResolved);
        statusDot.classList.toggle('incomplete', !allResolved);
    }

    // Update confirm button
    const confirmBtn = document.getElementById('mfrConfirmBtn');
    if (confirmBtn) {
        confirmBtn.disabled = !allResolved;
    }

    // Update panel border color when all resolved
    const panel = document.getElementById('mfrResolutionPanel');
    if (panel) {
        panel.classList.toggle('all-resolved', allResolved);
    }
}

/**
 * Cancel resolution - reject the promise and hide panel
 */
function cancelMfrResolution() {
    if (state.mfrResolutionPromise) {
        state.mfrResolutionPromise.reject(new Error('Resolution cancelled by user'));
    }
    hideMfrResolutionPanel();
    showStatus('Submission cancelled. Queue remains intact.', 'info');
}

/**
 * Confirm all resolutions - save to Supabase and resolve promise
 */
async function confirmMfrResolutions() {
    if (state.mfrResolutions.size !== state.unresolvedManufacturers.length) {
        showStatus('Please resolve all manufacturers before confirming', 'error');
        return;
    }

    // Disable button to prevent double-clicks
    const confirmBtn = document.getElementById('mfrConfirmBtn');
    if (confirmBtn) confirmBtn.disabled = true;

    showStatus('Saving manufacturer mappings...', 'loading');

    try {
        // Build the mapping list for batch save
        const mappingsToSave = [];
        const normalizedMap = new Map();

        for (const [index, resolution] of state.mfrResolutions) {
            const mfr = state.unresolvedManufacturers[index];
            const canonicalName = resolution.value;
            const source = mfr.distributor;

            // Add to mappings to save
            mappingsToSave.push({
                distributor_name: mfr.distributorName,
                canonical_name: canonicalName,
                distributor: source
            });

            // Add to normalized map for immediate use
            normalizedMap.set(mfr.distributorName, canonicalName);
        }

        // Save to Supabase
        const saveResult = await saveManufacturerMappingsBatch(mappingsToSave);

        if (saveResult.success || saveResult.saved_count >= 0) {
            console.log(`[MfrResolution] Saved ${saveResult.saved_count} mappings`);

            // Hide panel and resolve promise with the normalized map
            hideMfrResolutionPanel();

            if (state.mfrResolutionPromise) {
                state.mfrResolutionPromise.resolve(normalizedMap);
            }

            showStatus(`Saved ${saveResult.saved_count} manufacturer mapping(s). Continuing submission...`, 'success');
        } else {
            throw new Error('Failed to save mappings');
        }
    } catch (error) {
        console.error('[MfrResolution] Error confirming resolutions:', error);
        showStatus('Error saving mappings: ' + error.message, 'error');
        // Re-enable button on error
        if (confirmBtn) confirmBtn.disabled = false;
    }
}

/**
 * Escape HTML for safe rendering
 */
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/**
 * Convert string to Title Case (first letter of each word capitalized, rest lowercase)
 */
function toTitleCase(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/(?:^|\s)\S/g, char => char.toUpperCase());
}

// =====================================================
// MANUFACTURER MAPPINGS REFERENCE PANEL
// =====================================================

/**
 * Load manufacturer mappings from Supabase
 */
async function loadManufacturerMappings() {
    try {
        console.log('[MfrMappings] Loading manufacturer mappings from Supabase...');
        const response = await fetch(`${SUPABASE_URL}/rest/v1/manufacturer_mappings?select=*&order=canonical_name`, {
            headers: {
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to load mappings: ${response.status}`);
        }

        state.manufacturerMappingsData = await response.json();
        console.log(`[MfrMappings] Loaded ${state.manufacturerMappingsData.length} mappings`);
    } catch (error) {
        console.error('[MfrMappings] Error loading mappings:', error);
        state.manufacturerMappingsData = [];
    }
}

/**
 * Toggle the manufacturer mappings panel visibility
 */
function toggleMfrMappingsPanel() {
    const isBulk = state.searchMode === 'bulk';
    const panel = document.getElementById(isBulk ? 'bulkMfrMappingsPanel' : 'mfrMappingsPanel');
    const singleBtn = document.getElementById('mfrMappingsBtn');
    const bulkBtn = document.getElementById('bulkMfrMappingsBtn');

    if (!panel) return;

    const isVisible = panel.style.display !== 'none';

    if (isVisible) {
        panel.style.display = 'none';
        singleBtn?.classList.remove('active');
        bulkBtn?.classList.remove('active');
    } else {
        renderMfrMappingsTable(isBulk);
        panel.style.display = 'block';
        if (isBulk) {
            bulkBtn?.classList.add('active');
        } else {
            singleBtn?.classList.add('active');
        }
        panel.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

/**
 * Render the manufacturer mappings table
 */
function renderMfrMappingsTable(isBulk) {
    const tbody = document.getElementById(isBulk ? 'bulkMfrMappingsTableBody' : 'mfrMappingsTableBody');
    const countEl = document.getElementById(isBulk ? 'bulkMfrMappingsCount' : 'mfrMappingsCount');
    if (!tbody) return;

    const data = state.manufacturerMappingsData || [];
    if (countEl) countEl.textContent = data.length;

    function renderAliasCell(aliases, tagClass) {
        if (!aliases || aliases.length === 0) {
            return '<td><span class="mfr-mappings-no-alias">--</span></td>';
        }
        const tags = aliases.map(a => '<span class="mfr-alias-tag ' + tagClass + '">' + escapeHtml(a) + '</span>').join('');
        return '<td><div class="mfr-alias-tags">' + tags + '</div></td>';
    }

    tbody.innerHTML = data.map(function(mapping) {
        return '<tr>' +
            '<td><span style="font-weight:600">' + escapeHtml(mapping.canonical_name) + '</span></td>' +
            renderAliasCell(mapping.ingram_micro_aliases, 'mfr-alias-tag-ingram') +
            renderAliasCell(mapping.td_synnex_aliases, 'mfr-alias-tag-synnex') +
            renderAliasCell(mapping.adi_global_aliases, 'mfr-alias-tag-adi') +
            '</tr>';
    }).join('');
}

/**
 * Initialize vertical resize for mappings panel
 */
function initMfrMappingsResize() {
    const resizeHandle = document.getElementById('mfrMappingsResize');
    const container = document.getElementById('mfrMappingsTableContainer');

    if (!resizeHandle || !container) return;

    let isResizing = false;
    let startY = 0;
    let startHeight = 0;

    resizeHandle.addEventListener('mousedown', (e) => {
        isResizing = true;
        startY = e.clientY;
        startHeight = container.offsetHeight;
        document.body.style.cursor = 'ns-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        const deltaY = e.clientY - startY;
        const newHeight = Math.max(80, Math.min(400, startHeight + deltaY));
        container.style.maxHeight = newHeight + 'px';
    });

    document.addEventListener('mouseup', () => {
        if (isResizing) {
            isResizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });

    // Also init bulk mappings resize
    const bulkResizeHandle = document.getElementById('bulkMfrMappingsResize');
    const bulkContainer = document.getElementById('bulkMfrMappingsTableContainer');

    if (bulkResizeHandle && bulkContainer) {
        let isBulkResizing = false;
        let bulkStartY = 0;
        let bulkStartHeight = 0;

        bulkResizeHandle.addEventListener('mousedown', (e) => {
            isBulkResizing = true;
            bulkStartY = e.clientY;
            bulkStartHeight = bulkContainer.offsetHeight;
            document.body.style.cursor = 'ns-resize';
            document.body.style.userSelect = 'none';
            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (!isBulkResizing) return;
            const deltaY = e.clientY - bulkStartY;
            const newHeight = Math.max(80, Math.min(400, bulkStartHeight + deltaY));
            bulkContainer.style.maxHeight = newHeight + 'px';
        });

        document.addEventListener('mouseup', () => {
            if (isBulkResizing) {
                isBulkResizing = false;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
            }
        });
    }
}

// ============================================
// CUSTOMER DISCOUNT % — Queue Enhancement
// ============================================

// Clamp discount to integer 0-100, default 0
function clampDiscount(val) {
    var n = parseInt(val, 10);
    if (isNaN(n) || n < 0) return 0;
    if (n > 100) return 100;
    return Math.floor(n);
}

// Update per-row discount on input
function updateQueueItemDiscount(mpn, inputEl) {
    var queue = getActiveQueue();
    var product = queue.find(function(p) {
        return (p.ingramPartNumber || p.vendorPartNumber) === mpn;
    });
    if (!product) return;
    if (inputEl.value === '') return;
    var val = clampDiscount(inputEl.value);
    product.customerDiscount = val;
    inputEl.classList.toggle('has-value', val > 0);
}

// Finalize per-row discount on blur (default empty to 0)
function finalizeQueueItemDiscount(mpn, inputEl) {
    var queue = getActiveQueue();
    var product = queue.find(function(p) {
        return (p.ingramPartNumber || p.vendorPartNumber) === mpn;
    });
    if (!product) return;
    product.customerDiscount = clampDiscount(inputEl.value);
    inputEl.value = product.customerDiscount;
    inputEl.classList.toggle('has-value', product.customerDiscount > 0);
}

// Apply discount to ALL products in queue
function applyAllDiscount() {
    var input = document.getElementById('discountApplyAllInput');
    var val = clampDiscount(input.value);
    input.value = val;
    var queue = getActiveQueue();
    queue.forEach(function(p) { p.customerDiscount = val; });

    // Update all per-row inputs
    document.querySelectorAll('.queue-item-discount').forEach(function(el) {
        el.value = val;
        el.classList.toggle('has-value', val > 0);
        flashQueueItem(el.closest('.queue-item'));
    });

    // Sync mfr header inputs
    document.querySelectorAll('.mfr-discount-input').forEach(function(el) {
        el.value = val;
    });
}

// Apply discount to a single manufacturer group
function applyMfrDiscount(mfr) {
    var headerInput = document.querySelector('.mfr-discount-input[data-mfr="' + mfr.replace(/"/g, '\\"') + '"]');
    if (!headerInput) return;
    var val = clampDiscount(headerInput.value);
    headerInput.value = val;
    var queue = getActiveQueue();
    queue.forEach(function(p) {
        var pMfr = p.vendorName || p.manufacturer || 'Unknown';
        if (pMfr === mfr) {
            p.customerDiscount = val;
        }
    });

    // Update matching per-row inputs
    document.querySelectorAll('.queue-item-discount').forEach(function(el) {
        var itemMpn = el.getAttribute('data-mpn');
        var product = queue.find(function(p) {
            return (p.ingramPartNumber || p.vendorPartNumber) === itemMpn;
        });
        if (product) {
            var pMfr = product.vendorName || product.manufacturer || 'Unknown';
            if (pMfr === mfr) {
                el.value = val;
                el.classList.toggle('has-value', val > 0);
                flashQueueItem(el.closest('.queue-item'));
            }
        }
    });
}

// Auto-apply mfr discount on blur (only if non-zero)
function applyMfrDiscountIfNonZero(mfr) {
    var headerInput = document.querySelector('.mfr-discount-input[data-mfr="' + mfr.replace(/"/g, '\\"') + '"]');
    if (!headerInput) return;
    var val = clampDiscount(headerInput.value);
    headerInput.value = val;
    if (val > 0) {
        applyMfrDiscount(mfr);
    }
}

// Reset ALL discounts to 0
function clearAllDiscounts() {
    var input = document.getElementById('discountApplyAllInput');
    if (input) input.value = 0;
    var queue = getActiveQueue();
    queue.forEach(function(p) { p.customerDiscount = 0; });

    document.querySelectorAll('.queue-item-discount').forEach(function(el) {
        el.value = 0;
        el.classList.remove('has-value');
        flashQueueItem(el.closest('.queue-item'));
    });

    document.querySelectorAll('.mfr-discount-input').forEach(function(el) {
        el.value = 0;
    });
}

// Flash animation on queue item
function flashQueueItem(el) {
    if (!el) return;
    el.classList.remove('discount-flash');
    void el.offsetWidth;
    el.classList.add('discount-flash');
}

// Toggle discount field visibility (eye icon)
function toggleDiscountVisibility() {
    var panel = document.querySelector('.queue-panel');
    var toggleBtn = document.getElementById('discountVisibilityToggle');
    var eyeOpen = document.getElementById('discountEyeOpen');
    var eyeClosed = document.getElementById('discountEyeClosed');
    if (!panel || !toggleBtn) return;

    var isHidden = panel.classList.toggle('discount-fields-hidden');
    toggleBtn.classList.toggle('fields-hidden', isHidden);
    if (eyeOpen) eyeOpen.style.display = isHidden ? 'none' : '';
    if (eyeClosed) eyeClosed.style.display = isHidden ? '' : 'none';
}

// Generate column labels HTML for queue
function createQueueColumnLabelsHTML() {
    return '<div class="queue-column-labels">' +
        '<span class="queue-col-label queue-col-label-drag"></span>' +
        '<span class="queue-col-label queue-col-label-qty">QTY</span>' +
        '<span class="queue-col-label queue-col-label-part"></span>' +
        '<span class="queue-col-label queue-col-label-indicator"></span>' +
        '<span class="queue-col-label queue-col-label-price"></span>' +
        '<span class="queue-col-label queue-col-label-disc">CUST DISC %</span>' +
        '<span class="queue-col-label queue-col-label-pct-spacer"></span>' +
        '<span class="queue-col-label queue-col-label-remove"></span>' +
        '</div>';
}

async function submitQueue() {
    console.log('[SubmitQueue] Function called');

    if (getActiveQueue().length === 0) {
        showStatus('No products in queue', 'error');
        return;
    }

    showStatus('Checking manufacturer mappings...', 'loading');
    console.log(`[SubmitQueue] Processing ${getActiveQueue().length} products`);

    // Step 1: Extract unique manufacturers from queued products
    const uniqueManufacturers = new Map();
    for (const product of getActiveQueue()) {
        const mfr = product.vendorName || product.manufacturer;
        const distributor = product._source === 'tdsynnex' ? 'tdsynnex' : 'ingram';
        if (mfr && !uniqueManufacturers.has(mfr)) {
            uniqueManufacturers.set(mfr, distributor);
        }
    }
    console.log('[SubmitQueue] Unique manufacturers:', Array.from(uniqueManufacturers.keys()));

    // Step 2: Check mappings in batch via Supabase RPC
    const manufacturerList = Array.from(uniqueManufacturers.entries()).map(([name, distributor]) => ({
        name: name,
        distributor: distributor
    }));

    let mappingResults = [];
    try {
        mappingResults = await checkManufacturerMappingsBatch(manufacturerList);
    } catch (error) {
        console.error('[SubmitQueue] Error checking mappings:', error);
        showStatus('Error checking manufacturer mappings: ' + error.message, 'error');
        return;
    }

    // Step 3: Separate resolved vs. unresolved
    const normalizedMap = new Map();
    const unresolvedList = [];

    for (const result of mappingResults) {
        // RPC returns: distributor_name, distributor_source, found, canonical_name, suggested_name
        if (result.found) {
            normalizedMap.set(result.distributor_name, result.canonical_name);
        } else {
            unresolvedList.push({
                distributorName: result.distributor_name,
                distributor: result.distributor_source
            });
        }
    }

    console.log('[SubmitQueue] Resolved:', Object.fromEntries(normalizedMap));
    console.log('[SubmitQueue] Unresolved:', unresolvedList);

    // Step 4: If any unresolved, show resolution panel and wait
    if (unresolvedList.length > 0) {
        console.log('[SubmitQueue] Showing resolution panel for', unresolvedList.length, 'manufacturers');

        // Check if we have pre-fetched manufacturers
        if (state.prefetchedManufacturers.length === 0) {
            showStatus('Warning: No Zoho manufacturers available. You may only create new names.', 'info');
        }

        try {
            // Wait for user to resolve all manufacturers
            const userResolutions = await showMfrResolutionPanel(unresolvedList);

            // Merge user resolutions into normalizedMap
            for (const [distributorName, canonicalName] of userResolutions) {
                normalizedMap.set(distributorName, canonicalName);
            }

            console.log('[SubmitQueue] After resolution, normalizedMap:', Object.fromEntries(normalizedMap));
        } catch (error) {
            // User cancelled
            console.log('[SubmitQueue] Resolution cancelled by user');
            return;
        }
    }

    // Step 5: Format products with normalized manufacturers
    showStatus('Preparing products for submission...', 'loading');

    // Normalize bulk products to single-mode shape for formatting
    let productsToFormat = getActiveQueue();
    if (state.searchMode === 'bulk') {
        productsToFormat = productsToFormat.map(product => {
            const raw = product._rawRpcRow;
            const dist = product._source || state.currentDistributor;

            if (raw) {
                // Use existing mappers to get single-mode shape
                let mapped;
                switch (dist) {
                    case 'tdsynnex': mapped = mapTDSynnexProduct(raw); break;
                    case 'adi': mapped = mapADIGlobalProduct(raw); break;
                    default: mapped = {
                        ingramPartNumber: raw.ingram_part_number || product._fileVpn || product.mpn || '',
                        vendorPartNumber: raw.vendor_part_number || raw.manufacturer_part_number || '',
                        vendorName: raw.manufacturer || raw.vendor_name || '',
                        description: raw.description_line_1 || raw.description || '',
                        category: raw.category || '',
                        subCategory: raw.subcategory || '',
                        retailPrice: parseFloat(raw.retail_price) || 0,
                        pricingData: { pricing: { retailPrice: parseFloat(raw.retail_price) || 0, customerPrice: parseFloat(raw.customer_price) || null } },
                        upcCode: raw.upc || '',
                        productType: raw.im_product_type || '',
                        _source: 'ingram'
                    };
                }
                mapped.qty = product.qty || 1;
                    mapped.customerDiscount = product.customerDiscount || 0;
                mapped.resellerPrice = product.resellerPrice || null;
                // For mapped products, ensure resellerPrice is accessible for the formatter
                if (!mapped.resellerPrice && raw) {
                    mapped.resellerPrice = parseFloat(raw.contract_price || raw.reseller_price || raw.customer_price || raw.unit_cost) || null;
                }
                // Overlay spreadsheet MSRP so it flows through to Zoho
                if (product.msrp !== null && product.msrp !== undefined) {
                    mapped.retailPrice = product.msrp;
                    if (mapped.pricingData && mapped.pricingData.pricing) {
                        mapped.pricingData.pricing.retailPrice = product.msrp;
                    }
                }
                return mapped;
            } else {
                // Fallback: build minimal shape from bulk fields
                return {
                    vendorPartNumber: product.mpn || product.vendorPartNumber || '',
                    vendorName: product.manufacturer || '',
                    description: product.description || '',
                    retailPrice: product.msrp || 0,
                    pricingData: { pricing: { retailPrice: product.msrp, customerPrice: product.resellerPrice } },
                    resellerPrice: product.resellerPrice || null,
                    _source: product._source || state.currentDistributor,
                    qty: product.qty || 1,
                        customerDiscount: product.customerDiscount || 0,
                    ...(dist === 'ingram' && { ingramPartNumber: product._fileVpn || product.vpn || product.mpn || '' }),
                    ...(dist === 'tdsynnex' && { tdSynnexSkuNumber: product.vpn || '', distributorPartNumber: product.vpn || '' }),
                    ...(dist === 'adi' && { adiSku: product.vpn || '', distributorPartNumber: product.vpn || '' }),
                };
            }
        });
    }

    const productsWithMissingMfr = [];

    const formattedProducts = productsToFormat.map(product => {
        const pricingData = product.pricingData || state.pricingData?.[product.ingramPartNumber] || {};
        const msrp = pricingData?.pricing?.retailPrice || product.retailPrice || null;
        const originalMfr = product.vendorName || product.manufacturer;
        const normalizedMfr = normalizedMap.get(originalMfr) || originalMfr;

        // Defensive check: Track products with missing manufacturer
        if (!normalizedMfr) {
            const sku = product.vendorPartNumber || product.ingramPartNumber || 'Unknown SKU';
            console.error(`[SubmitQueue] Missing manufacturer for product: ${sku}`, {
                vendorName: product.vendorName,
                stateManufacturer: state.manufacturer,
                originalMfr,
                normalizedMfr,
                product
            });
            productsWithMissingMfr.push(sku);
        }

        // TD Synnex format
        if (product._source === 'tdsynnex') {
            return {
                Product_Code: product.vendorPartNumber || '',
                Product_Name: product.description || '',
                Manufacturer: normalizedMfr,
                TDSynnex_SKU: product.distributorPartNumber || '',
                Replacement_SKU: (product.replacementSku && product.replacementSku.trim() && !/^[\r\n]+$/.test(product.replacementSku)) ? product.replacementSku.trim() : 'None',
                MSRP: msrp,
                // Customer_Price MUST match what the queue displays as reseller price — no transformation
                Customer_Price: product.resellerPrice || pricingData?.pricing?.customerPrice || product.contractPrice || product.unitCost || null,
                Category_Level_1: product.category || state.category || '',
                Category_Level_2: product.subCategory || state.subcategory || '',
                Category_Level_3: product.cat3 || state.cat3 || '',
                UPC: product.upcCode || '',
                Description: product.extraDescription || '',
                Last_Sync_Source: 'TD SYNNEX',
                UNSPSC_Commodity: product.commodityName || '',
                Kit_or_Standalone: product.kitStandaloneFlag === 'K' ? 'Yes' : 'No',
                Quantity: product.qty || 1,
                        Customer_Discount: parseInt(product.customerDiscount) || 0
            };
        }

        // ADI Global format
        if (product._source === 'adi') {
            return {
                Product_Code: product.vendorPartNumber || '',
                Product_Name: product.description || '',
                Manufacturer: normalizedMfr,
                ADI_SKU: product.adiSku || product.distributorPartNumber || '',
                MSRP: msrp,
                Customer_Price: product.resellerPrice || pricingData?.pricing?.customerPrice || null,
                ADI_Category_1: product.category || '',
                ADI_Category_2: product.category2 || '',
                UPC: product.upcCode || '',
                Description: product.extraDescription || '',
                Last_Sync_Source: 'ADI Global',
                Quantity: product.qty || 1,
                Customer_Discount: parseInt(product.customerDiscount) || 0
            };
        }

        // Ingram Micro format (default)
        return {
            Product_Code: product.vendorPartNumber || '',
            Product_Name: product.description || '',
            Manufacturer: normalizedMfr,
            Ingram_Micro_SKU: product.ingramPartNumber || '',
            Replacement_SKU: (product.replacementSku && product.replacementSku.trim() && !/^[\r\n]+$/.test(product.replacementSku)) ? product.replacementSku.trim() : 'None',
            MSRP: msrp,
            // Customer_Price MUST match what the queue displays as reseller price — no transformation
            Customer_Price: product.resellerPrice || pricingData?.pricing?.customerPrice || null,
            Category: product.category || state.category || '',
            Subcategory: product.subCategory || state.subcategory || '',
            UPC: pricingData?.upc || product.upcCode || '',
            Description: product.extraDescription || pricingData?.description || '',
            Last_Sync_Source: 'Ingram Micro',
            IM_Product_Type: product.productType || '',
            Kit_or_Standalone: pricingData?.bundlePartIndicator ? 'Yes' : 'No',
            Quantity: product.qty || 1,
                        Customer_Discount: parseInt(product.customerDiscount) || 0
        };
    });

    console.log('[SubmitQueue] Formatted products:', formattedProducts);
    console.log('[SubmitQueue] Manufacturer values:', formattedProducts.map(p => p.Manufacturer));

    // Prevent submission if any products have missing manufacturer
    if (productsWithMissingMfr.length > 0) {
        const errorMsg = `Cannot submit: Missing manufacturer for SKU(s): ${productsWithMissingMfr.join(', ')}. Please try re-searching for these products.`;
        console.error('[SubmitQueue] Aborting submission due to missing manufacturers:', productsWithMissingMfr);
        showStatus(errorMsg, 'error');
        return;
    }

    // Step 6: Submit to Zoho
    if (typeof $Client !== 'undefined') {
        console.log('[SubmitQueue] Calling $Client.close...');
        $Client.close({
            products: formattedProducts,
            distributor: state.currentDistributor
        });
    } else {
        console.log('[SubmitQueue] Standalone mode - would send:', formattedProducts);
        showStatus(`Queued ${formattedProducts.length} products (standalone mode)`, 'info');
    }
}

// =====================================================
// BATCH PRICING (fallback)
// =====================================================
async function fetchBatchPricing(products) {
    const partNumbers = products
        .map(p => p.ingramPartNumber)
        .filter(pn => pn);

    if (partNumbers.length === 0) return;

    try {
        const response = await fetch(`${PROXY_BASE}?action=pricing`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ partNumbers, sandbox: false })
        });

        const data = await response.json();
        state.pricingData = {};

        if (Array.isArray(data)) {
            data.forEach(item => {
                state.pricingData[item.ingramPartNumber] = item;
            });
        }
    } catch (error) {
        console.error('[Pricing] Error:', error);
    }
}

// =====================================================
// PRODUCT DETAILS
// =====================================================
async function showProductDetails(productIndex) {
    const product = state.currentProducts[productIndex];
    if (!product) {
        console.error('Product not found at index:', productIndex);
        return;
    }

    const isTDSynnex = product._source === 'tdsynnex';
    const isADI = product._source === 'adi';
    const distLabel = isTDSynnex ? 'TD Synnex' : isADI ? 'ADI Global' : 'Ingram';
    console.log(`[Details] Loading details for ${product.vendorPartNumber} (${distLabel})...`);

    const detailsSection = document.getElementById('productDetailsSection');
    detailsSection.style.display = 'block';
    detailsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

    // Reset raw API visibility
    state.rawApiVisible = false;
    const rawContainer = document.getElementById('rawApiContainer');
    const rawToggle = document.getElementById('rawApiToggle');
    if (rawContainer) rawContainer.style.display = 'none';
    if (rawToggle) rawToggle.classList.remove('active');

    // Clear previous data and show loading indicator
    document.getElementById('detailsProductName').innerHTML = '';
    document.getElementById('detailsSubtitle').innerHTML = '';
    var longDescClear = document.getElementById('detailsLongDesc');
    if (longDescClear) longDescClear.style.display = 'none';
    ['productInfoGrid', 'pricingGrid', 'availabilityGrid', 'flagsGrid'].forEach(function(id) {
        var el = document.getElementById(id);
        if (el) el.innerHTML = '';
    });
    var discG = document.getElementById('discountsGroup');
    if (discG) discG.style.display = 'none';
    var whSec = document.getElementById('warehouseSection');
    if (whSec) whSec.style.display = 'none';
    document.getElementById('rawApiResponse').textContent = '';

    // Show loading indicator inside details content
    var detailsContent = document.querySelector('#productDetailsSection .details-content');
    var loadingEl = document.getElementById('detailsLoadingIndicator');
    if (!loadingEl && detailsContent) {
        loadingEl = document.createElement('div');
        loadingEl.id = 'detailsLoadingIndicator';
        loadingEl.style.cssText = 'display:flex;align-items:center;justify-content:center;gap:10px;padding:32px 16px;color:var(--color-text-secondary);font-size:var(--font-size-sm);';
        loadingEl.innerHTML = '<div style="width:20px;height:20px;border:2.5px solid var(--color-border);border-top-color:var(--color-accent);border-radius:50%;animation:spin 0.8s linear infinite;flex-shrink:0;"></div>Loading realtime pricing and inventory data';
        detailsContent.insertBefore(loadingEl, detailsContent.querySelector('.details-section'));
    }
    if (loadingEl) loadingEl.style.display = 'flex';

    // Helper functions (shared between distributors)
    const yesNo = (val) => {
        if (val === true) return 'Yes';
        if (val === false) return 'No';
        if (typeof val === 'string') {
            const lower = val.toLowerCase();
            if (lower === 'true' || lower === 'yes') return 'Yes';
            if (lower === 'false' || lower === 'no') return 'No';
        }
        return '-';
    };

    const formatCurrency = (val) => {
        if (val === null || val === undefined) return '-';
        return `$${Number(val).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    };

    const renderGrid = (elementId, fields) => {
        const grid = document.getElementById(elementId);
        if (grid) {
            grid.innerHTML = fields.map(f => `
                <div class="field-mapping-item">
                    <span class="field-label">${f.label}</span>
                    <span class="field-value">${f.value}</span>
                </div>
            `).join('');
        }
    };

    const renderGridWithOptions = (elementId, fields) => {
        const grid = document.getElementById(elementId);
        if (grid) {
            grid.innerHTML = fields.map(f => `
                <div class="field-mapping-item${f.fullWidth ? ' full-width' : ''}">
                    <span class="field-label">${f.label}</span>
                    <span class="field-value">${f.value}</span>
                </div>
            `).join('');
        }
    };

    const renderFlagsGrid = (elementId, fields) => {
        const grid = document.getElementById(elementId);
        if (grid) {
            grid.innerHTML = fields.map(f => `
                <div class="field-mapping-item">
                    <span class="field-label">${f.label}</span>
                    <span class="field-value">${f.isHtml ? f.value : f.value}</span>
                </div>
            `).join('');
        }
    };

    // ========================================
    // TD SYNNEX PRODUCT DETAILS
    // ========================================
    if (isTDSynnex) {
        // Access raw product data for TD SYNNEX fields
        const rawProduct = product._rawProduct || {};
        let warehouseData = { warehouses: [], totalQty: 0, totalOnOrder: 0, status: null };

        // Fetch warehouse availability from TD SYNNEX XML API using numeric SKU (Field 5)
        if (product.tdSynnexSkuNumber) {
            warehouseData = await fetchTDSynnexWarehouseAvailability(product.tdSynnexSkuNumber);
        }

        const fullProductData = { ...product, warehouseData };

        // Determine authorization status from API response
        const isNotAuthorized = warehouseData.status === 'Notauthorized';
        const authorizedText = isNotAuthorized ? 'No' : 'Yes';
        const authorizedClass = isNotAuthorized ? 'authorized-no' : 'authorized-yes';

        // Header - Product Name (from part_description)
        document.getElementById('detailsProductName').innerHTML = `
            <strong>Product Name:</strong> ${product.description || 'N/A'}
        `;

        // Header - TD Synnex SKU (Field 5), Vendor Part (manufacturer_part_number), Manufacturer, Authorized
        document.getElementById('detailsSubtitle').innerHTML = `
            <strong>TD Synnex SKU:</strong> ${product.tdSynnexSkuNumber || 'N/A'} |
            <strong>Vendor Part:</strong> ${product.vendorPartNumber || 'N/A'} |
            <strong>Manufacturer:</strong> ${product.vendorName || state.manufacturer} |
            <strong>Authorized:</strong> <span class="${authorizedClass}">${authorizedText}</span>
        `;

        // Long Description (from long_description)
        const longDesc = product.extraDescription || '';
        const longDescEl = document.getElementById('detailsLongDesc');
        if (longDesc) {
            longDescEl.innerHTML = `<strong>Long Description:</strong> ${longDesc}`;
            longDescEl.style.display = 'block';
        } else {
            longDescEl.style.display = 'none';
        }

        // Product Information Grid - TD SYNNEX specific labels
        // Category 1 = cat_description_1, Category 2 = cat_description_2, Category 3 = cat_description_3
        // SKU Type = derived from Fields 28,53,54,55 (weight/dimensions)
        // UNSPSC = commodity_name
        // Replacement SKU = replacement_sku
        const productInfoFields = [
            { label: 'Category 1', value: product.category || '-' },
            { label: 'Category 2', value: product.subCategory || '-' },
            { label: 'Category 3', value: product.cat3 || '-' },
            { label: 'SKU Type', value: product.skuType || '-' },
            { label: 'UNSPSC', value: product.commodityName || '-' },
            { label: 'Replacement SKU', value: product.replacementSku || '-' }
        ];
        renderGrid('productInfoGrid', productInfoFields);

        // Pricing Grid - TD SYNNEX: MSRP from msrp, Customer Price from contract_price
        // Using pricingData.pricing which is set in mapTDSynnexProduct from raw fields
        const pricingFields = [
            { label: 'MSRP', value: formatCurrency(product.pricingData?.pricing?.retailPrice) },
            { label: 'Customer Price', value: formatCurrency(product.pricingData?.pricing?.customerPrice) },
            { label: 'Regular Price', value: formatCurrency(product.unitCost) }
        ];
        renderGrid('pricingGrid', pricingFields);

        // Availability Grid - use API response if available, else flat file data (qty_total)
        const apiTotalQty = warehouseData.totalQty;
        const flatFileTotalQty = product.pricingData?.availability?.totalAvailability ?? 0;
        const displayQty = apiTotalQty > 0 ? apiTotalQty : flatFileTotalQty;
        const inStock = displayQty > 0;

        const availabilityFields = [
            { label: 'In Stock', value: yesNo(inStock) },
            { label: 'Available Qty', value: displayQty === 9999 ? 'Unlimited' : displayQty }
        ];
        renderGrid('availabilityGrid', availabilityFields);

        // Flags Grid - TD SYNNEX specific derivation
        // Digital = SKU Type derived from weight/dimensions
        // Bundle = kit_standalone_flag = "K"
        // Licensed = td_assigned_use (Field 36) contains "License"
        // Service SKU = cat_description_1 = "Service / Support" OR cat_description_2 contains "Service"/"Support"
        // Direct Ship = sku_attributes first char is "Y"
        // New = sku_created_date (Field 37) <= 90 days
        // Discontinued = abc_code = "C" or "T"
        const discontinuedValue = product.isDiscontinued
            ? '<span class="discontinued-yes">Yes</span>'
            : '<span class="discontinued-no">No</span>';

        const flagsFields = [
            { label: 'Digital', value: yesNo(product.isDigital) },
            { label: 'Bundle', value: yesNo(product.kitStandaloneFlag === 'K') },
            { label: 'Licensed', value: yesNo(product.isLicensed) },
            { label: 'Service SKU', value: yesNo(product.isServiceSku) },
            { label: 'Direct Ship', value: yesNo(product.skuAttributes?.charAt(0) === 'Y') },
            { label: 'New', value: yesNo(product.isNew) },
            { label: 'Discontinued', value: discontinuedValue, isHtml: true }
        ];
        renderFlagsGrid('flagsGrid', flagsFields);

        // Discounts - TD SYNNEX: show only if promo_flag = "Y"
        // Type = "Rebate" (static), Bid Number = "N/A" (static)
        // Discount = unit_cost - contract_price
        // Qty = 99999 (static), Effective = "N/A" (static), Expires = promo_expiration
        const discountsGroup = document.getElementById('discountsGroup');
        const discountsBody = document.getElementById('discountsBody');

        if (product.promoFlag === 'Y') {
            discountsGroup.style.display = 'block';
            // Discount = unit_cost - contract_price (from raw fields via mapped properties)
            const discountAmount = (product.unitCost && product.contractPrice)
                ? product.unitCost - product.contractPrice
                : null;
            discountsBody.innerHTML = `
                <tr>
                    <td>Rebate</td>
                    <td>N/A</td>
                    <td class="text-right">${formatCurrency(discountAmount)}</td>
                    <td class="text-right">99999</td>
                    <td>N/A</td>
                    <td>${formatPromoDate(product.promoExpiration)}</td>
                </tr>
            `;
        } else {
            discountsGroup.style.display = 'none';
        }

        // Warehouse Availability - from TD SYNNEX XML API response
        // Warehouse = WHS-001 (number), Location = WHS-003 (city)
        // Available = WHS-005 (qty), Backordered = WHS-006 (onOrderQuantity)
        const warehouseSection = document.getElementById('warehouseSection');
        const warehouseBody = document.getElementById('warehouseBody');

        // Only show warehouse card when In Stock (qty > 0)
        const availableWarehouses = (warehouseData.warehouses || []).filter(wh => (wh.qty ?? 0) > 0);

        if (availableWarehouses.length > 0) {
            warehouseSection.style.display = 'block';
            warehouseBody.innerHTML = availableWarehouses.map(wh => `
                <tr>
                    <td>${wh.warehouseId || wh.number || '-'}</td>
                    <td>${wh.city || '-'}</td>
                    <td class="text-right">${wh.qty ?? 0}</td>
                    <td class="text-right">${wh.onOrder ?? wh.onOrderQuantity ?? 0}</td>
                </tr>
            `).join('');
        } else {
            warehouseSection.style.display = 'none';
        }

        document.getElementById('rawApiResponse').textContent = JSON.stringify(fullProductData, null, 2);
        var _li = document.getElementById('detailsLoadingIndicator');
        if (_li) _li.style.display = 'none';
        scrollToPanel('productDetailsSection');
        return;
    }

    // ========================================
    // ADI GLOBAL PRODUCT DETAILS
    // ========================================
    if (isADI) {
        // Fetch live pricing/inventory from ADI proxy using VPN (item/adiSku)
        const adiVpn = product.adiSku || product.distributorPartNumber || '';
        let adiApiData = null;
        let adiItemData = null;

        if (adiVpn) {
            try {
                const adiResponse = await fetch(`${ADI_PROXY_BASE}?action=pricing&vpns=${encodeURIComponent(adiVpn)}`);
                if (adiResponse.ok) {
                    adiApiData = await adiResponse.json();
                    // Extract item from ItemList array
                    if (adiApiData?.ItemList && Array.isArray(adiApiData.ItemList) && adiApiData.ItemList.length > 0) {
                        adiItemData = adiApiData.ItemList[0];
                    }
                }
            } catch (err) {
                console.error('[ADI Details] Error fetching pricing:', err);
            }
        }

        const fullProductData = { ...product, adiApiResponse: adiApiData };

        // Determine authorization from API response (AllowedToBuy: "Y"/"N")
        const isAuthorized = adiItemData?.AllowedToBuy === 'Y';
        const authorizedText = isAuthorized ? 'Yes' : 'No';
        const authorizedClass = isAuthorized ? 'authorized-yes' : 'authorized-no';

        // Header - Product Name
        document.getElementById('detailsProductName').innerHTML = `
            <strong>Product Name:</strong> ${product.description || 'N/A'}
        `;

        // Header - ADI SKU, Vendor Part, Manufacturer, Authorized
        document.getElementById('detailsSubtitle').innerHTML = `
            <strong>ADI SKU:</strong> ${adiVpn || 'N/A'} |
            <strong>Vendor Part:</strong> ${product.vendorPartNumber || 'N/A'} |
            <strong>Manufacturer:</strong> ${product.vendorName || state.manufacturer} |
            <strong>Authorized:</strong> <span class="${authorizedClass}">${authorizedText}</span>
        `;

        // Long Description
        const longDesc = product.extraDescription || '';
        const longDescEl = document.getElementById('detailsLongDesc');
        if (longDesc) {
            longDescEl.innerHTML = `<strong>Long Description:</strong> ${longDesc}`;
            longDescEl.style.display = 'block';
        } else {
            longDescEl.style.display = 'none';
        }

        // Product Information Grid - ADI specific
        const productInfoFields = [
            { label: 'Category 1', value: product.category || '-' },
            { label: 'Category 2', value: product.category2 || '-' },
            { label: 'UPC', value: product.upcCode || '-' }
        ];
        renderGrid('productInfoGrid', productInfoFields);

        // Pricing Grid - ADI: use live API price if available, fall back to DB
        const adiLivePrice = adiItemData?.ItemPrice ? parseFloat(adiItemData.ItemPrice) : null;
        const adiMsrp = product.pricingData?.pricing?.retailPrice;
        const adiCustomerPrice = adiLivePrice || product.pricingData?.pricing?.customerPrice;
        const pricingFields = [
            { label: 'MSRP', value: formatCurrency(adiMsrp) },
            { label: 'Customer Price', value: formatCurrency(adiCustomerPrice) }
        ];
        renderGrid('pricingGrid', pricingFields);

        // Availability from API — NationalInventory is a string number
        const nationalInventoryStr = adiItemData?.NationalInventory;
        const nationalInventory = nationalInventoryStr ? parseInt(nationalInventoryStr, 10) : 0;
        const inStock = nationalInventory > 0;
        const availabilityFields = [
            { label: 'In Stock', value: yesNo(inStock) },
            { label: 'Available Qty', value: nationalInventory || 0 }
        ];
        renderGrid('availabilityGrid', availabilityFields);

        // No flags for ADI
        renderFlagsGrid('flagsGrid', []);

        // No discounts section for ADI
        const discountsGroup = document.getElementById('discountsGroup');
        if (discountsGroup) discountsGroup.style.display = 'none';

        // No warehouse section for ADI
        const warehouseSection = document.getElementById('warehouseSection');
        if (warehouseSection) warehouseSection.style.display = 'none';

        document.getElementById('rawApiResponse').textContent = JSON.stringify(fullProductData, null, 2);
        var _li = document.getElementById('detailsLoadingIndicator');
        if (_li) _li.style.display = 'none';
        scrollToPanel('productDetailsSection');
        return;
    }

    // ========================================
    // INGRAM MICRO PRODUCT DETAILS
    // ========================================
    const ingramPn = product.ingramPartNumber;
    let pricingData = state.pricingData?.[ingramPn];
    let productDetails = null;

    if (ingramPn) {
        const fetchPromises = [];

        if (!pricingData) {
            fetchPromises.push(
                fetch(`${PROXY_BASE}?action=pricing`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ partNumbers: [ingramPn], sandbox: false })
                })
                .then(res => res.json())
                .then(data => {
                    if (Array.isArray(data) && data.length > 0) {
                        pricingData = data[0];
                        state.pricingData[ingramPn] = pricingData;
                    }
                })
                .catch(err => console.error('[Details] Error fetching pricing:', err))
            );
        }

        fetchPromises.push(
            fetch(`${PROXY_BASE}?action=productDetails&ingramPartNumber=${encodeURIComponent(ingramPn)}`)
                .then(res => res.json())
                .then(data => {
                    if (data && !data.error) {
                        productDetails = data;
                    }
                })
                .catch(err => console.error('[Details] Error fetching product details:', err))
        );

        await Promise.all(fetchPromises);
    }

    const fullProductData = { ...product, pricingData, productDetails };

    const isAuthorized = product.authorizedToPurchase === 'true' ||
                         product.authorizedToPurchase === true ||
                         pricingData?.productAuthorized === true;
    const authorizedText = isAuthorized ? 'Yes' : 'No';
    const authorizedClass = isAuthorized ? 'authorized-yes' : 'authorized-no';

    // Row 1: Product Name
    document.getElementById('detailsProductName').innerHTML = `
        <strong>Product Name:</strong> ${product.description || 'N/A'}
    `;
    // Row 2: Ingram SKU, Vendor Part, Manufacturer, Authorized
    document.getElementById('detailsSubtitle').innerHTML = `
        <strong>Ingram SKU:</strong> ${ingramPn || 'N/A'} |
        <strong>Vendor Part:</strong> ${product.vendorPartNumber || 'N/A'} |
        <strong>Manufacturer:</strong> ${product.vendorName || state.manufacturer} |
        <strong>Authorized:</strong> <span class="${authorizedClass}">${authorizedText}</span>
    `;

    const longDesc = product.extraDescription || pricingData?.description || '';
    const longDescEl = document.getElementById('detailsLongDesc');
    if (longDesc) {
        longDescEl.innerHTML = `<strong>Long Description:</strong> ${longDesc}`;
        longDescEl.style.display = 'block';
    } else {
        longDescEl.style.display = 'none';
    }

    // Ingram: Category, Subcategory, Product Type, SKU Type, Product Class, Replacement SKU
    const productInfoFields = [
        { label: 'Category', value: product.category || state.category || '-' },
        { label: 'Subcategory', value: product.subCategory || state.subcategory || '-' },
        { label: 'Product Type', value: product.productType || '-' },
        { label: 'Media Type', value: product.type || product.productType || '-' },
        { label: 'Product Class', value: formatProductClass(pricingData?.productClass || product.productClass), fullWidth: true },
        { label: 'Replacement SKU', value: product.replacementSku || '-', fullWidth: true }
    ];
    renderGridWithOptions('productInfoGrid', productInfoFields);

    // Ingram pricing: MSRP (retailPrice) and Customer Price (customerPrice) - no Subscription Price
    const msrpValue = formatCurrency(pricingData?.pricing?.retailPrice);
    const customerPriceValue = formatCurrency(pricingData?.pricing?.customerPrice);

    // Calculate Regular Price = Customer Price + sum of all discounts
    const customerPrice = pricingData?.pricing?.customerPrice || 0;
    let totalDiscounts = 0;
    if (pricingData?.discounts && Array.isArray(pricingData.discounts)) {
        pricingData.discounts.forEach(discountGroup => {
            if (discountGroup.specialPricing && Array.isArray(discountGroup.specialPricing)) {
                discountGroup.specialPricing.forEach(sp => {
                    if (sp.specialPricingDiscount) {
                        totalDiscounts += parseFloat(sp.specialPricingDiscount) || 0;
                    }
                });
            }
        });
    }
    const regularPrice = customerPrice + totalDiscounts;

    const pricingFields = [
        { label: 'MSRP', value: msrpValue },
        { label: 'Customer Price', value: customerPriceValue },
        { label: 'Regular Price', value: formatCurrency(regularPrice) }
    ];
    renderGrid('pricingGrid', pricingFields);

    // Ingram discounts
    const discountsGroup = document.getElementById('discountsGroup');
    const discountsBody = document.getElementById('discountsBody');

    let allDiscounts = [];
    if (pricingData?.discounts && Array.isArray(pricingData.discounts)) {
        pricingData.discounts.forEach(discountGroup => {
            if (discountGroup.specialPricing && Array.isArray(discountGroup.specialPricing)) {
                allDiscounts.push(...discountGroup.specialPricing);
            }
        });
    }

    if (allDiscounts.length > 0) {
        discountsGroup.style.display = 'block';
        discountsBody.innerHTML = allDiscounts.map(d => `
            <tr>
                <td>${d.discountType || '-'}</td>
                <td>${d.specialBidNumber || '-'}</td>
                <td class="text-right">${formatCurrency(d.specialPricingDiscount)}</td>
                <td class="text-right">${d.specialPricingAvailableQuantity ?? '-'}</td>
                <td>${d.specialPricingEffectiveDate || '-'}</td>
                <td>${d.specialPricingExpirationDate || '-'}</td>
            </tr>
        `).join('');
    } else {
        discountsGroup.style.display = 'none';
    }

    const availabilityFields = [
        { label: 'In Stock', value: yesNo(pricingData?.availability?.available) },
        { label: 'Available Qty', value: pricingData?.availability?.totalAvailability ?? '-' }
    ];
    renderGrid('availabilityGrid', availabilityFields);

    const indicators = productDetails?.indicators || {};

    // Discontinued badge with color
    const isDiscontinued = product.discontinued || indicators.isDiscontinuedProduct;
    const discontinuedValue = isDiscontinued === true || isDiscontinued === 'true'
        ? '<span class="discontinued-yes">Yes</span>'
        : '<span class="discontinued-no">No</span>';

    // Ingram flags
    const flagsFields = [
        { label: 'Digital', value: yesNo(indicators.isDigitalType || product.type === 'IM::Digital' || product.type === 'IM::digital') },
        { label: 'Bundle', value: yesNo(indicators.hasBundle || pricingData?.bundlePartIndicator) },
        { label: 'Licensed', value: yesNo(indicators.isLicenseProduct) },
        { label: 'Service SKU', value: yesNo(indicators.isServiceSku) },
        { label: 'Direct Ship', value: yesNo(product.directShip || indicators.isDirectship) },
        { label: 'New', value: yesNo(product.newProduct || indicators.isNewProduct) },
        { label: 'Discontinued', value: discontinuedValue, isHtml: true }
    ];
    renderFlagsGrid('flagsGrid', flagsFields);

    // Ingram warehouse availability
    const warehouseSection = document.getElementById('warehouseSection');
    const warehouseBody = document.getElementById('warehouseBody');

    const totalAvailability = pricingData?.availability?.totalAvailability ?? 0;

    if (totalAvailability > 0 && pricingData?.availability?.availabilityByWarehouse?.length > 0) {
        const availableWarehouses = pricingData.availability.availabilityByWarehouse
            .filter(wh => (wh.quantityAvailable ?? 0) > 0);

        if (availableWarehouses.length > 0) {
            warehouseSection.style.display = 'block';
            warehouseBody.innerHTML = availableWarehouses.map(wh => `
                <tr>
                    <td>${wh.warehouseId}</td>
                    <td>${wh.location || '-'}</td>
                    <td class="text-right">${wh.quantityAvailable ?? 0}</td>
                    <td class="text-right">${wh.quantityBackordered ?? 0}</td>
                </tr>
            `).join('');
        } else {
            warehouseSection.style.display = 'none';
        }
    } else {
        warehouseSection.style.display = 'none';
    }

    document.getElementById('rawApiResponse').textContent = JSON.stringify(fullProductData, null, 2);
    var _li = document.getElementById('detailsLoadingIndicator');
    if (_li) _li.style.display = 'none';
    scrollToPanel('productDetailsSection');
}

function hideProductDetails() {
    document.getElementById('productDetailsSection').style.display = 'none';
}

// =====================================================
// ACTION HANDLERS (Legacy support)
// =====================================================
function addSelectedProducts() {
    // Legacy function - now redirects to queue workflow
    addSelectedToQueue();
}

function closeWidget() {
    if (typeof $Client !== 'undefined') {
        $Client.close({ cancelled: true, products: [] });
    }
}

function cancelSelection() {
    console.log('Cancel clicked');

    if (typeof $Client !== 'undefined') {
        $Client.close({ cancelled: true, products: [] });
    }

    state.selectedProducts.clear();
    state.queuedProducts = [];
    bulkState.queuedProducts = [];
    updateSelectedCount();
    updateQueueUI();
}

// =====================================================
// UTILITY FUNCTIONS
// =====================================================
function resetFilters() {
    state.manufacturer = '';
    state.currentPage = 1;
    state.totalRecords = 0;
    state.totalPages = 1;

    // Reset SKU search mode
    state.skuSearchMode = false;
    state.pendingSkuFilter = '';
    state.skuManufacturerOptions = [];

    document.getElementById('manufacturerSearch').value = '';
    if (state.currentDistributor === 'ingram') {
        // Ingram: re-populate the pre-loaded manufacturer dropdown
        loadIngramManufacturers();
    } else if (state.currentDistributor === 'tdsynnex') {
        // TD Synnex: re-populate the pre-loaded manufacturer dropdown
        loadTDSynnexManufacturers();
    } else {
        document.getElementById('manufacturerSelect').innerHTML =
            '<option value="">Type to search manufacturers...</option>';
    }
    document.getElementById('mfrCount').textContent = '';
    document.getElementById('selectedMfrBadge').textContent = '';

    // Re-apply distributor-specific mode class on manufacturer combo
    const mfrComboEl = document.querySelector('.mfr-combo');
    if (mfrComboEl) {
        mfrComboEl.classList.remove('ingram-mode', 'adi-mode', 'tdsynnex-mode');
        if (state.currentDistributor === 'ingram') {
            mfrComboEl.classList.add('ingram-mode');
        } else if (state.currentDistributor === 'adi') {
            mfrComboEl.classList.add('adi-mode');
        } else if (state.currentDistributor === 'tdsynnex') {
            mfrComboEl.classList.add('tdsynnex-mode');
        }
    }

    // Clear SKU search input and reset placeholder
    const skuSearch = document.getElementById('skuSearch');
    const skuSearchRow = document.getElementById('skuSearchRow');
    if (skuSearch) {
        skuSearch.value = '';
        skuSearch.placeholder = 'Enter partial or full SKU (e.g. AB123, XYZ-456)...';
    }
    if (skuSearchRow) {
        skuSearchRow.classList.remove('filter-mode');
    }

    // Show OR divider
    const orDivider = document.getElementById('orDivider');
    if (orDivider) {
        orDivider.classList.remove('hidden');
    }

    document.getElementById('optionalFiltersRow').style.display = 'none';

    resetOptionalFilters();
    resetProducts();

    document.getElementById('productsSection').style.display = 'none';
    showStatus('Select a manufacturer or enter a SKU to begin', 'info');

    // Close Manufacturer Mappings panel if open
    const mappingsPanel = document.getElementById('mfrMappingsPanel');
    const mappingsBtn = document.getElementById('mfrMappingsBtn');
    if (mappingsPanel) mappingsPanel.style.display = 'none';
    if (mappingsBtn) mappingsBtn.classList.remove('active');

    // Close Resolve Manufacturer Names panel if open
    hideMfrResolutionPanel();
}

function resetOptionalFilters() {
    state.category = '';
    state.subcategory = '';
    state.cat3 = '';
    state.skuType = '';
    state.skuKeyword = '';

    state.filterParams.category = '';
    state.filterParams.subcategory = '';
    state.filterParams.cat3 = '';
    state.filterParams.skuType = '';

    const catSelect = document.getElementById('categorySelect');
    if (catSelect) {
        catSelect.innerHTML = '<option value="">-- Any --</option>';
        document.getElementById('catCount').textContent = '';
    }

    const subSelect = document.getElementById('subcategorySelect');
    if (subSelect) {
        subSelect.innerHTML = '<option value="">-- Any --</option>';
        document.getElementById('subCatCount').textContent = '';
    }

    const cat3Select = document.getElementById('cat3Select');
    if (cat3Select) {
        cat3Select.innerHTML = '<option value="">-- Any --</option>';
        document.getElementById('cat3Count').textContent = '';
    }

    const skuTypeSelect = document.getElementById('skuTypeSelect');
    if (skuTypeSelect) {
        skuTypeSelect.innerHTML = '<option value="">-- Any --</option>';
    }
    const skuTypeCount = document.getElementById('skuTypeCount');
    if (skuTypeCount) skuTypeCount.textContent = '';

    const skuSearch = document.getElementById('skuSearch');
    if (skuSearch) {
        skuSearch.value = '';
    }
}

function resetProducts() {
    document.getElementById('productsBody').innerHTML = '';
    document.getElementById('pagination').innerHTML = '';
    document.getElementById('productCount').textContent = '0 products';
    document.getElementById('productDetailsSection').style.display = 'none';

    // Only clear current page selection, NOT the queue
    state.selectedProducts.clear();
    state.currentProducts = [];
    state.pricingData = {};

    updateSelectedCount();

    // Reset select all checkbox
    const selectAll = document.getElementById('selectAll');
    if (selectAll) selectAll.checked = false;
}

let statusTimeout = null;

function showStatus(message, type) {
    const el = document.getElementById('filterStatus');
    if (!el) return;

    // Clear any existing timeout
    if (statusTimeout) {
        clearTimeout(statusTimeout);
        statusTimeout = null;
    }

    el.className = `status-bar ${type}`;
    el.innerHTML = `<span class="status-message">${message}</span>`;

    if (!message) {
        el.style.display = 'none';
    } else {
        el.style.display = 'flex';

        // Auto-dismiss success and info messages after 10 seconds
        if (type === 'success' || type === 'info') {
            statusTimeout = setTimeout(() => {
                el.style.display = 'none';
            }, 10000);
        }
    }
}

// =====================================================
// BULK SEARCH — STATE & MODE TOGGLE (Phase 1)
// Completely independent from single-search code above.
// =====================================================

const bulkState = {
    initialized: false,
    detectedDistributor: null,
    detectionSucceeded: false,
    workbook: null,
    fileRows: [],
    parsedSkus: [],
    fileName: null,
    // Phase 3
    parsedFileData: null,
    selectionMode: null,
    hiddenColumns: new Set(),
    hiddenRows: new Set(),
    previewZoom: 55,
    userManuallyZoomed: false,
    userHasResized: false,
    activeHiddenRowsDropdown: null,
    lastDataColumn: 0, // 0-based index of last column with actual data
    products: [],
    unmatchedMpns: [],
    isLoading: false,
    rawRpcRows: new Map(),
    // Phase 5
    selectedProductIndices: new Set(),
    collapsedGroups: new Set(),
    pricingMode: 'reseller',
    // Phase 6 — isolated queue
    queuedProducts: [],
    // Phase 7b — MSRP comparison
    msrpMismatches: [],
    msrpChoices: new Map(),
    resultsPricingMode: 'reseller',
    // Phase 8 — pagination
    resultsPage: 1,
    resultsPerPage: 50,
};

const BULK_PREVIEW_MAX_ROWS = 50;

function setSearchMode(mode) {
    if (mode === state.searchMode) return;
    state.searchMode = mode;

    const singlePanel = document.querySelector('.single-search-panel');
    const bulkPanel = document.querySelector('.bulk-search-panel-container');
    // Update bulk toggle button active state
    const bulkToggle = document.getElementById('bulkToggle');
    if (bulkToggle) {
        bulkToggle.classList.toggle('active', mode === 'bulk');
    }

    // Close product details panel when switching modes
    hideProductDetails();

    // Close manufacturer mappings panel and deactivate both buttons
    const mfrPanel = document.getElementById('mfrMappingsPanel');
    if (mfrPanel) mfrPanel.style.display = 'none';
    const bulkMfrPanel = document.getElementById('bulkMfrMappingsPanel');
    if (bulkMfrPanel) bulkMfrPanel.style.display = 'none';
    document.getElementById('mfrMappingsBtn')?.classList.remove('active');
    document.getElementById('bulkMfrMappingsBtn')?.classList.remove('active');

    // Queue panel (#rightPanel) lives inside .single-search-panel > .panels-row.
    // To keep it visible in bulk mode, we hide single-mode children selectively
    // instead of hiding the entire .single-search-panel.
    const contentWrapper = document.querySelector('.content-wrapper');

    if (mode === 'single') {
        // Restore single-search-panel: show always-visible children, leave conditionally-hidden ones alone
        singlePanel.querySelectorAll('.left-panel').forEach(el => el.style.display = '');
        singlePanel.style.display = '';
        bulkPanel.style.display = 'none';
        // Restore normal layout
        contentWrapper.classList.remove('bulk-mode-active');
        singlePanel.classList.remove('queue-only');
        // Single mode: restore default distributor (full UI sync via selectDistributor)
        selectDistributor('ingram');
    } else {
        // Hide single-mode content but keep queue (.right-panel) visible
        singlePanel.querySelectorAll('.left-panel').forEach(el => el.style.display = 'none');
        // Hide any lingering status notification from single mode
        const filterStatus = document.getElementById('filterStatus');
        if (filterStatus) filterStatus.style.display = 'none';
        bulkPanel.style.display = '';
        // Activate flex row layout so bulk content + queue sit side by side
        contentWrapper.classList.add('bulk-mode-active');
        singlePanel.classList.add('queue-only');
        // Lazy init drop zone on first bulk activation
        if (!bulkState.initialized) {
            bulkState.initialized = true;
            bulkInitDropZone();
            document.addEventListener('click', bulkHandleDropdownOutsideClick);
        }
        // Bulk mode: no default distributor — user must select or auto-detect
        state.currentDistributor = null;
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        // Update distributor badges AFTER null reset so stale single-mode value is cleared
        updateBulkDistributorBadges();
        bulkUpdateLoadButtonState();
    }

    // Sync pricing toggle to active mode's pricing state
    const activePricing = getActivePricingMode();
    document.querySelectorAll('.pricing-toggle-btn').forEach(btn =>
        btn.classList.toggle('active', btn.dataset.price === activePricing)
    );
    // Hide pricing toggle in single mode (MSRP only)
    const pricingToggle = document.getElementById('pricingToggle');
    if (pricingToggle) {
        pricingToggle.style.display = (mode === 'single') ? 'none' : '';
    }
            // Customer Discount: hidden by default in single mode, visible in bulk mode
            var queuePanel = document.querySelector('.queue-panel');
            var discountToggleBtn = document.getElementById('discountVisibilityToggle');
            var eyeOpen = document.getElementById('discountEyeOpen');
            var eyeClosed = document.getElementById('discountEyeClosed');
            if (queuePanel) {
                if (mode === 'single') {
                    queuePanel.classList.add('discount-fields-hidden');
                    if (discountToggleBtn) discountToggleBtn.classList.add('fields-hidden');
                    if (eyeOpen) eyeOpen.style.display = 'none';
                    if (eyeClosed) eyeClosed.style.display = '';
                } else {
                    queuePanel.classList.remove('discount-fields-hidden');
                    if (discountToggleBtn) discountToggleBtn.classList.remove('fields-hidden');
                    if (eyeOpen) eyeOpen.style.display = '';
                    if (eyeClosed) eyeClosed.style.display = 'none';
                }
            }
    // Re-render queue for active mode
    updateQueueUI();
    updateFooterStats();
}

function handleBulkToggleClick() {
    if (state.searchMode === 'bulk') {
        setSearchMode('single');
        document.getElementById('bulkToggle').classList.remove('active');
    } else {
        setSearchMode('bulk');
        document.getElementById('bulkToggle').classList.add('active');
    }
}

function bulkUpdateLoadButtonState() {
    const loadBtn = document.getElementById('bulkLoadProductsBtn');
    if (!loadBtn) return;

    const hasDistributor = !!state.currentDistributor;
    const hasPasteContent = !!(document.getElementById('bulkPasteArea') && document.getElementById('bulkPasteArea').value.trim());
    const hasFile = !!(bulkState.fileRows && bulkState.fileRows.length > 0);
    const hasInput = hasFile || hasPasteContent;

    let shouldDisable = false;
    let pulseButtons = false;
    let statusMsg = null;

    if (!hasInput) {
        // No file or paste — disable but no pulse (user hasn't done anything yet)
        shouldDisable = true;
    } else if (hasFile && bulkState.detectionSucceeded && hasDistributor && bulkState.detectedDistributor !== state.currentDistributor) {
        // File mode: detection succeeded but user selected WRONG distributor
        shouldDisable = true;
        pulseButtons = true;
        const detectedName = DISTRIBUTORS[bulkState.detectedDistributor]?.name || bulkState.detectedDistributor;
        const selectedName = DISTRIBUTORS[state.currentDistributor]?.name || state.currentDistributor;
        statusMsg = { text: 'Detected ' + detectedName + ' in file but ' + selectedName + ' is selected. Please select ' + detectedName + ' to proceed.', type: 'warning' };
    } else if (hasFile && !bulkState.detectionSucceeded && !hasDistributor) {
        // File mode: detection failed AND no distributor selected
        shouldDisable = true;
        pulseButtons = true;
        statusMsg = { text: 'Could not auto-detect distributor from file. You must choose a distributor manually to proceed.', type: 'warning' };
    } else if (hasPasteContent && !hasDistributor) {
        // Paste mode: no distributor selected
        shouldDisable = true;
        pulseButtons = true;
        statusMsg = { text: 'Please select a distributor to proceed.', type: 'warning' };
    } else if (!hasDistributor) {
        // Has input but no distributor — disable, no pulse (shouldn't normally reach here)
        shouldDisable = true;
    }

    loadBtn.disabled = shouldDisable;

    // Pulse animation on distributor buttons
    const tabsContainer = document.querySelector('.distributor-tabs');
    if (tabsContainer) {
        if (pulseButtons) {
            tabsContainer.classList.add('pulse-attention');
        } else {
            tabsContainer.classList.remove('pulse-attention');
        }
    }

    // Show persistent status message if needed
    if (statusMsg) {
        showStatus(statusMsg.text, statusMsg.type);
    }
}

// =====================================================
// BULK SEARCH — FILE & PASTE HANDLING (Phase 2d)
// =====================================================

/**
 * Initialize drag/drop event listeners on the bulk drop zone.
 * Called once lazily when bulk mode is first activated.
 */
function bulkInitDropZone() {
    const dropZone = document.getElementById('bulkDropZone');
    if (!dropZone) return;

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'copy';
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const fileInput = document.getElementById('bulkFileInput');
            if (fileInput) {
                fileInput.files = files;
                bulkHandleFileSelect({ target: fileInput });
            }
        }
    });

    console.log('[BulkSearch] Drop zone initialized');
}

/**
 * Handle file selection for bulk search (file input or drag/drop).
 * Reads the file with SheetJS and stores the workbook in bulkState.
 */
function bulkHandleFileSelect(event) {
    const file = event.target.files ? event.target.files[0] : null;
    if (!file) return;

    const dropZone = document.getElementById('bulkDropZone');
    const statusEl = document.getElementById('bulkDropZoneStatus');
    const spinner = document.getElementById('bulkLoadingSpinner');

    if (statusEl) {
        statusEl.textContent = `Processing ${file.name}...`;
        statusEl.style.color = 'var(--color-warning)';
    }

    // Show loading spinner
    if (spinner) spinner.classList.add('visible');

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            bulkState.workbook = workbook;
            bulkState.fileName = file.name;

            // Sheet selector (Phase 3 elements — guard since they don't exist yet)
            const el = document.getElementById('bulkSheetSelect');
            if (el) {
                if (workbook.SheetNames.length > 1) {
                    el.innerHTML = '';
                    workbook.SheetNames.forEach((name, index) => {
                        const option = document.createElement('option');
                        option.value = index;
                        option.textContent = name;
                        el.appendChild(option);
                    });
                    el.classList.add('visible');
                    const label = document.getElementById('bulkSheetSelectLabel');
                    if (label) label.classList.add('visible');
                } else {
                    el.classList.remove('visible');
                    const label = document.getElementById('bulkSheetSelectLabel');
                    if (label) label.classList.remove('visible');
                }
            }

            // Load first sheet by default
            bulkLoadSheetData(0);

            // Update drop zone UI on success
            if (dropZone) dropZone.classList.add('has-file');
            const clearBtn = document.getElementById('bulkDropZoneClearBtn');
            if (clearBtn) clearBtn.classList.add('visible');
            if (statusEl) {
                statusEl.textContent = `${file.name} loaded`;
                statusEl.style.color = 'var(--color-success)';
            }

            console.log(`[BulkSearch] Loaded ${workbook.SheetNames.length} sheet(s) from ${file.name}`);
        } catch (err) {
            console.error('[BulkSearch] Parse error:', err);
            if (statusEl) {
                statusEl.textContent = `Error: ${err.message}`;
                statusEl.style.color = 'var(--color-error)';
            }
        } finally {
            // Hide spinner regardless of outcome
            if (spinner) spinner.classList.remove('visible');
        }
    };

    reader.onerror = function() {
        if (statusEl) {
            statusEl.textContent = 'Error reading file';
            statusEl.style.color = 'var(--color-error)';
        }
        if (spinner) spinner.classList.remove('visible');
    };

    reader.readAsArrayBuffer(file);
}

/**
 * Load data from a specific sheet in the bulk workbook.
 * Stores raw 2D array in bulkState.fileRows.
 */
function bulkLoadSheetData(sheetIndex) {
    if (!bulkState.workbook) return;

    const sheetName = bulkState.workbook.SheetNames[sheetIndex];
    const worksheet = bulkState.workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (rows.length === 0) {
        console.warn('[BulkSearch] Empty sheet:', sheetName);
        bulkState.fileRows = [];
        bulkHideSpreadsheetPreview();
        return;
    }

    bulkState.fileRows = rows;
    console.log(`[BulkSearch] Loaded ${rows.length} rows from sheet: ${sheetName}`);

    // Auto-detect distributor from spreadsheet content (Phase 9.2 Step 2)
    console.log('[BulkDetect] About to call bulkDetectDistributor(), fileRows length:', bulkState.fileRows.length);
    bulkDetectDistributor();

    // Phase 3: Enable mappings panel and render preview
    const mappingsPanel = document.getElementById('bulkMappingsPanel');
    if (mappingsPanel) mappingsPanel.classList.remove('disabled');

    // Populate column dropdowns from row 0 (default header row 1) without auto-selecting
    bulkUpdateColumnSelectionDropdown(0, null, true);

    // Render preview immediately
    bulkState.userManuallyZoomed = false;
    bulkState.previewZoom = 55;
    bulkRenderSpreadsheetPreview();

    showStatus('File loaded \u2014 use Auto Map or select columns manually', 'info');
}

/**
 * Auto Map button handler: detects header row, fetches mapping rules,
 * auto-maps columns, and detects last data row.
 */
function bulkAutoMapColumns() {
    if (!bulkState.fileRows || bulkState.fileRows.length === 0) {
        showStatus('Load a file first', 'warning');
        return;
    }

    // Re-detect distributor
    bulkDetectDistributor();

    // Auto-detect header row (async), then fetch rules, then map columns
    bulkAutoDetectHeaderRow().then(function(detectedHeaderRow) {
        var headerRowInput = document.getElementById('bulkHeaderRowInput');
        if (headerRowInput) {
            headerRowInput.value = detectedHeaderRow + 1; // 1-based for UI
        }
        console.log('[BulkAutoMap] Setting header row to', detectedHeaderRow + 1);

        return bulkFetchMappingRules().then(function(rules) {
            bulkUpdateColumnSelectionDropdown(detectedHeaderRow, rules);

            var lastRow = bulkAutoDetectLastRow(detectedHeaderRow);
            var bulkLastRowInput = document.getElementById('bulkLastRowInput');
            if (bulkLastRowInput && lastRow > detectedHeaderRow) {
                bulkLastRowInput.value = lastRow + 1;
                console.log('[BulkAutoMap] Auto-detected last data row:', lastRow + 1);
            }
        });
    }).then(function() {
        bulkState.userManuallyZoomed = false;
        bulkState.previewZoom = 55;
        bulkRenderSpreadsheetPreview();

        var headerRowInput = document.getElementById('bulkHeaderRowInput');
        var lastRowInput = document.getElementById('bulkLastRowInput');
        var detectedRow = headerRowInput ? parseInt(headerRowInput.value) : 1;
        var detectedLastRow = lastRowInput ? lastRowInput.value : '';
        var statusMsg = 'Auto-mapped: header row ' + detectedRow;
        if (detectedLastRow) statusMsg += ', last row ' + detectedLastRow;
        showStatus(statusMsg, 'info');
    });
}

// =====================================================
// BULK SEARCH — Phase 9.2+9.4: Mapping Rules Editor
// =====================================================

var _bulkMappingRulesResizeInited = false;

// Column definitions for the mapping rules editor
var BULK_RULES_COLUMNS = [
    { key: 'distributor', label: 'Distributor', collapsible: false, editable: false },
    { key: 'sheet_name', label: 'Sheet', collapsible: true, editable: true },
    { key: 'mpn', label: 'MPN', collapsible: true, editable: true },
    { key: 'qty', label: 'QTY', collapsible: true, editable: true },
    { key: 'price', label: 'Price', collapsible: true, editable: true },
    { key: 'vpn', label: 'VPN', collapsible: true, editable: true },
    { key: 'msrp', label: 'MSRP', collapsible: true, editable: true }
];

var BULK_RULES_DIST_NAMES = { ingram: 'Ingram Micro', tdsynnex: 'TD Synnex', adiglobal: 'ADI Global' };

// Track collapsed columns and current edit cell
var _bulkRuleCollapsedCols = {};
var _bulkRulesCurrentEditCell = null;

function bulkToggleMappingRulesPanel() {
    var panel = document.getElementById('bulkMappingRulesPanel');
    var btn = document.getElementById('bulkMappingRulesBtn');
    if (!panel) return;

    var isVisible = panel.style.display !== 'none';

    if (isVisible) {
        panel.style.display = 'none';
        if (btn) btn.classList.remove('active');
    } else {
        bulkFetchAllMappingRules();
        panel.style.display = 'block';
        if (btn) btn.classList.add('active');
        if (!_bulkMappingRulesResizeInited) {
            initBulkMappingRulesResize();
            _bulkMappingRulesResizeInited = true;
        }
        document.getElementById('bulkMappingRulesPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function initBulkMappingRulesResize() {
    var handle = document.getElementById('bulkMappingRulesResize');
    var wrap = document.getElementById('bulkMappingRulesTableContainer');
    if (!handle || !wrap) return;

    var isResizing = false;
    var startY = 0;
    var startH = 0;

    handle.addEventListener('mousedown', function(e) {
        isResizing = true;
        startY = e.clientY;
        startH = wrap.offsetHeight;
        document.body.style.cursor = 'ns-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    });

    document.addEventListener('mousemove', function(e) {
        if (!isResizing) return;
        var delta = e.clientY - startY;
        var newH = Math.max(100, Math.min(600, startH + delta));
        wrap.style.maxHeight = newH + 'px';
    });

    document.addEventListener('mouseup', function() {
        if (isResizing) {
            isResizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });

    console.log('[BulkRulesEditor] Resize handle initialized');
}

function bulkFetchAllMappingRules() {
    var url = SUPABASE_URL + '/rest/v1/bulk_column_mapping_rules?select=distributor,sheet_name,mpn,qty,price,vpn,msrp&order=distributor.asc';

    fetch(url, {
        headers: {
            'apikey': SUPABASE_ANON_KEY,
            'Authorization': 'Bearer ' + SUPABASE_ANON_KEY
        }
    })
    .then(function(res) { return res.json(); })
    .then(function(rows) {
        console.log('[BulkRulesEditor] Fetched', rows.length, 'mapping rules');
        bulkRenderMappingRulesTable(rows);
    })
    .catch(function(err) {
        console.error('[BulkRulesEditor] Failed to fetch rules:', err);
        var tbody = document.getElementById('bulkMappingRulesTableBody');
        if (tbody) {
            tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; color: var(--color-text-muted); padding: 12px;">Failed to load mapping rules</td></tr>';
        }
    });
}

function bulkRenderMappingRulesTable(rows) {
    if (!rows || rows.length === 0) {
        var tbody = document.getElementById('bulkMappingRulesTableBody');
        if (tbody) {
            tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; color: var(--color-text-muted); padding: 12px;">No mapping rules found</td></tr>';
        }
        return;
    }

    bulkRenderMappingRulesHeader();
    bulkRenderMappingRulesBody(rows);
    bulkSetupRuleColumnResizeHandles();
}

function bulkRenderMappingRulesHeader() {
    var headerRow = document.getElementById('bulkMappingRulesHeaderRow');
    if (!headerRow) return;
    var html = '';

    for (var i = 0; i < BULK_RULES_COLUMNS.length; i++) {
        var col = BULK_RULES_COLUMNS[i];
        var isCollapsed = !!_bulkRuleCollapsedCols[i];
        var colClass = isCollapsed ? ' bulk-rules-col-collapsed' : '';

        html += '<th class="' + colClass.trim() + '" data-col="' + i + '">';
        html += '<span class="bulk-rules-col-label">' + col.label + '</span>';

        if (col.collapsible) {
            html += '<button class="bulk-rules-col-toggle" data-col="' + i + '" title="' + (isCollapsed ? 'Expand' : 'Collapse') + ' column">';
            html += '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 9l-7 7-7-7"/></svg>';
            html += '</button>';
        }

        // Resize handle (skip last column)
        if (i < BULK_RULES_COLUMNS.length - 1) {
            html += '<div class="bulk-rules-col-resize" data-col="' + i + '"></div>';
        }

        html += '</th>';
    }

    headerRow.innerHTML = html;

    // Bind toggle events on chevron buttons
    var toggleBtns = headerRow.querySelectorAll('.bulk-rules-col-toggle');
    for (var t = 0; t < toggleBtns.length; t++) {
        (function(btn) {
            btn.addEventListener('click', function(e) {
                e.stopPropagation();
                var colIdx = parseInt(btn.getAttribute('data-col'), 10);
                bulkToggleRuleColumn(colIdx);
            });
        })(toggleBtns[t]);
    }

    // Bind click-to-expand on collapsed column headers
    var allThs = headerRow.querySelectorAll('th[data-col]');
    for (var h = 0; h < allThs.length; h++) {
        (function(th) {
            th.addEventListener('click', function() {
                var colIdx = parseInt(th.getAttribute('data-col'), 10);
                if (_bulkRuleCollapsedCols[colIdx]) {
                    bulkToggleRuleColumn(colIdx);
                }
            });
        })(allThs[h]);
    }
}

function bulkRenderMappingRulesBody(data) {
    var tbody = document.getElementById('bulkMappingRulesTableBody');
    if (!tbody) return;
    var html = '';

    for (var r = 0; r < data.length; r++) {
        var row = data[r];
        var isUniversal = row.distributor === 'universal';
        var rowClass = isUniversal ? ' class="bulk-rules-universal-row"' : '';

        html += '<tr' + rowClass + ' data-distributor="' + escapeHtml(row.distributor) + '">';

        // Distributor cell
        if (isUniversal) {
            html += '<td><span class="bulk-rules-universal-label">Universal</span>';
            html += '<span class="bulk-rules-universal-sub">(fallback)</span></td>';
        } else {
            var name = BULK_RULES_DIST_NAMES[row.distributor] || row.distributor;
            html += '<td>' + escapeHtml(name) + '</td>';
        }

        // Data cells
        for (var c = 1; c < BULK_RULES_COLUMNS.length; c++) {
            var col = BULK_RULES_COLUMNS[c];
            var value = row[col.key] || '';
            var isCollapsed = !!_bulkRuleCollapsedCols[c];
            var tdClass = 'bulk-rules-editable';
            if (isCollapsed) tdClass += ' bulk-rules-col-collapsed';

            html += '<td class="' + tdClass + '" data-col="' + c + '" data-field="' + col.key + '" data-distributor="' + escapeHtml(row.distributor) + '">';

            // Edit icon
            html += '<svg class="bulk-rules-edit-icon" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">';
            html += '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>';
            html += '<path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>';

            // Content
            html += '<span class="bulk-rules-cell-content">';
            if (value) {
                html += bulkRenderKeywordTags(value);
            } else {
                html += '<span class="bulk-rules-cell-empty">(none)</span>';
            }
            html += '</span>';

            // Collapsed indicator
            html += '<span class="bulk-rules-cell-collapsed-indicator ' + (value ? 'has-data' : 'no-data') + '"></span>';

            html += '</td>';
        }

        html += '</tr>';
    }

    tbody.innerHTML = html;

    // Bind click-to-edit on editable cells
    var editCells = tbody.querySelectorAll('td.bulk-rules-editable');
    for (var e = 0; e < editCells.length; e++) {
        (function(cell) {
            cell.addEventListener('click', function() {
                if (cell.classList.contains('bulk-rules-col-collapsed')) return;
                bulkEditMappingRuleCell(cell);
            });
        })(editCells[e]);
    }
}

function bulkRenderKeywordTags(value) {
    var parts = value.split(',');
    var html = '<span class="bulk-rules-keywords">';
    for (var i = 0; i < parts.length; i++) {
        var kw = parts[i].trim();
        if (kw) {
            html += '<span class="bulk-rules-keyword-tag">' + escapeHtml(kw) + '</span>';
        }
    }
    html += '</span>';
    return html;
}

function bulkEditMappingRuleCell(td) {
    // If already editing another cell, save it first
    if (_bulkRulesCurrentEditCell && _bulkRulesCurrentEditCell !== td) {
        bulkFinishEditingRuleCell(_bulkRulesCurrentEditCell, true);
    }
    if (td.classList.contains('bulk-rules-editing')) return;

    var field = td.getAttribute('data-field');
    var distributor = td.getAttribute('data-distributor');

    // Extract current raw value from tags
    var tags = td.querySelectorAll('.bulk-rules-keyword-tag');
    var currentValue = '';
    if (tags.length > 0) {
        var vals = [];
        for (var i = 0; i < tags.length; i++) {
            vals.push(tags[i].textContent);
        }
        currentValue = vals.join(', ');
    }

    td.classList.add('bulk-rules-editing');
    _bulkRulesCurrentEditCell = td;

    var textarea = document.createElement('textarea');
    textarea.className = 'bulk-rules-edit-textarea';
    textarea.value = currentValue;
    textarea.placeholder = 'e.g. keyword1, keyword2';
    textarea.setAttribute('data-original', currentValue);

    var hint = document.createElement('div');
    hint.className = 'bulk-rules-edit-hint';
    hint.innerHTML = '<span><kbd>Enter</kbd> save &middot; <kbd>Esc</kbd> cancel</span>';

    // Clear cell and insert editor
    td.innerHTML = '';
    td.appendChild(textarea);
    td.appendChild(hint);

    // Auto-size
    bulkAutoResizeRuleTextarea(textarea);
    textarea.focus();
    textarea.select();

    // Events
    textarea.addEventListener('input', function() {
        bulkAutoResizeRuleTextarea(textarea);
    });

    textarea.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            bulkFinishEditingRuleCell(td, true);
        }
        if (e.key === 'Escape') {
            e.preventDefault();
            bulkFinishEditingRuleCell(td, false);
        }
    });

    textarea.addEventListener('blur', function() {
        // Small delay to allow button clicks
        setTimeout(function() {
            if (td.classList.contains('bulk-rules-editing')) {
                bulkFinishEditingRuleCell(td, true);
            }
        }, 150);
    });
}

function bulkAutoResizeRuleTextarea(textarea) {
    textarea.style.height = 'auto';
    textarea.style.height = Math.max(52, textarea.scrollHeight) + 'px';
}

function bulkFinishEditingRuleCell(td, shouldSave) {
    var textarea = td.querySelector('.bulk-rules-edit-textarea');
    if (!textarea) return;

    var field = td.getAttribute('data-field');
    var distributor = td.getAttribute('data-distributor');
    var originalValue = textarea.getAttribute('data-original') || '';
    var newValue = textarea.value.trim().toLowerCase();

    // Normalize: trim each keyword, remove empty, deduplicate
    if (newValue) {
        var parts = newValue.split(',');
        var seen = {};
        var cleaned = [];
        for (var i = 0; i < parts.length; i++) {
            var kw = parts[i].trim();
            if (kw && !seen[kw]) {
                seen[kw] = true;
                cleaned.push(kw);
            }
        }
        newValue = cleaned.join(',');
    }

    // Build normalized original for comparison
    var normalizedOriginal = originalValue.split(',').map(function(k) { return k.trim(); }).filter(function(k) { return k; }).join(',');

    if (!shouldSave) {
        // Revert to original
        newValue = normalizedOriginal;
    }

    td.classList.remove('bulk-rules-editing');
    _bulkRulesCurrentEditCell = null;

    // Rebuild cell content
    bulkRestoreMappingRuleCell(td, distributor, field, newValue);

    if (shouldSave && newValue !== normalizedOriginal) {
        bulkSaveMappingRuleField(distributor, field, newValue, td);
    }
}

function bulkSaveMappingRuleField(distributor, field, newValue, td) {
    var patchUrl = SUPABASE_URL + '/rest/v1/bulk_column_mapping_rules?distributor=eq.' +
        encodeURIComponent(distributor);

    var patchBody = { updated_at: new Date().toISOString() };
    patchBody[field] = newValue;

    fetch(patchUrl, {
        method: 'PATCH',
        headers: {
            'apikey': SUPABASE_ANON_KEY,
            'Authorization': 'Bearer ' + SUPABASE_ANON_KEY,
            'Content-Type': 'application/json',
            'Prefer': 'return=minimal'
        },
        body: JSON.stringify(patchBody)
    })
    .then(function(res) {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        console.log('[BulkRulesEditor] Saved ' + field + ' for ' + distributor + ':', newValue);
    })
    .catch(function(err) {
        console.error('[BulkRulesEditor] Failed to save:', err);
        showStatus('Failed to save mapping rule', 'error');
    });
}

function bulkRestoreMappingRuleCell(td, distributor, field, value) {
    var html = '';
    html += '<svg class="bulk-rules-edit-icon" width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">';
    html += '<path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>';
    html += '<path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>';
    html += '<span class="bulk-rules-cell-content">';
    if (value) {
        html += bulkRenderKeywordTags(value);
    } else {
        html += '<span class="bulk-rules-cell-empty">(none)</span>';
    }
    html += '</span>';
    html += '<span class="bulk-rules-cell-collapsed-indicator ' + (value ? 'has-data' : 'no-data') + '"></span>';
    td.innerHTML = html;

    // Re-bind click
    td.addEventListener('click', function() {
        if (!td.classList.contains('bulk-rules-col-collapsed')) {
            bulkEditMappingRuleCell(td);
        }
    });
}

function bulkToggleRuleColumn(colIndex) {
    _bulkRuleCollapsedCols[colIndex] = !_bulkRuleCollapsedCols[colIndex];
    console.log('[BulkRulesEditor] Column ' + colIndex + ' collapsed:', _bulkRuleCollapsedCols[colIndex]);
    bulkApplyCollapsedCols();
}

function bulkApplyCollapsedCols() {
    var table = document.querySelector('.bulk-rules-table');
    if (!table) return;

    // Update colgroup
    var cols = document.querySelectorAll('#bulkMappingRulesColgroup col');
    var expandedCount = 0;
    for (var i = 0; i < BULK_RULES_COLUMNS.length; i++) {
        if (i === 0) continue; // distributor col stays fixed
        if (!_bulkRuleCollapsedCols[i]) expandedCount++;
    }

    for (var j = 0; j < cols.length; j++) {
        if (j === 0) {
            cols[j].style.width = '100px';
        } else if (_bulkRuleCollapsedCols[j]) {
            cols[j].style.width = '32px';
        } else {
            // Distribute remaining space equally among expanded columns
            cols[j].style.width = (1 / expandedCount * 100) + '%';
        }
    }

    // Update header
    var ths = table.querySelectorAll('thead th');
    for (var h = 0; h < ths.length; h++) {
        var isCollapsed = !!_bulkRuleCollapsedCols[h];
        if (isCollapsed) {
            ths[h].classList.add('bulk-rules-col-collapsed');
        } else {
            ths[h].classList.remove('bulk-rules-col-collapsed');
        }
        var toggle = ths[h].querySelector('.bulk-rules-col-toggle');
        if (toggle) {
            toggle.title = isCollapsed ? 'Expand column' : 'Collapse column';
        }
    }

    // Update body cells
    var tds = table.querySelectorAll('tbody td[data-col]');
    for (var d = 0; d < tds.length; d++) {
        var ci = parseInt(tds[d].getAttribute('data-col'), 10);
        if (_bulkRuleCollapsedCols[ci]) {
            tds[d].classList.add('bulk-rules-col-collapsed');
        } else {
            tds[d].classList.remove('bulk-rules-col-collapsed');
        }
    }
}

function bulkSetupRuleColumnResizeHandles() {
    var handles = document.querySelectorAll('.bulk-rules-col-resize');
    for (var i = 0; i < handles.length; i++) {
        (function(handle) {
            var colIdx = parseInt(handle.getAttribute('data-col'), 10);
            var isDragging = false;
            var startX = 0;
            var startWidth = 0;
            var colEl = null;

            handle.addEventListener('mousedown', function(e) {
                isDragging = true;
                startX = e.clientX;
                colEl = document.querySelectorAll('#bulkMappingRulesColgroup col')[colIdx];
                var th = handle.parentElement;
                startWidth = th.offsetWidth;
                document.body.style.cursor = 'col-resize';
                document.body.style.userSelect = 'none';
                e.preventDefault();
                e.stopPropagation();
            });

            document.addEventListener('mousemove', function(e) {
                if (!isDragging) return;
                var delta = e.clientX - startX;
                var newW = Math.max(40, startWidth + delta);
                if (colEl) colEl.style.width = newW + 'px';
            });

            document.addEventListener('mouseup', function() {
                if (isDragging) {
                    isDragging = false;
                    document.body.style.cursor = '';
                    document.body.style.userSelect = '';
                }
            });
        })(handles[i]);
    }
}

/**
 * Parse SKUs from the bulk paste textarea.
 * Deduplicates and normalizes (spaces→#).
 * Stores result in bulkState.parsedSkus.
 */
function bulkParsePastedSKUs() {
    const textarea = document.getElementById('bulkPasteArea');
    if (!textarea) return;

    const text = textarea.value.trim();

    if (!text) {
        bulkState.parsedSkus = [];
        bulkUpdateParsedPreview();
        bulkUpdateLoadButtonState();
        updateFooterStats();
        console.log('[BulkSearch] Parsed 0 SKUs from paste (empty input)');
        return;
    }

    const skus = text
        .split(/\r?\n/)
        .filter(line => line.trim().length > 0)
        .flatMap(line => {
            const normalized = line.trim().replace(/\s+/g, '#');
            // Split on comma or tab (separators within a line)
            return normalized.split(/[,\t]+/);
        })
        .map(s => s.trim().toUpperCase())
        .filter(s => s.length > 0);

    // Deduplicate
    bulkState.parsedSkus = [...new Set(skus)];
    console.log(`[BulkSearch] Parsed ${bulkState.parsedSkus.length} SKUs from paste`);
    bulkUpdateParsedPreview();
    bulkUpdateLoadButtonState();
    updateFooterStats();
    scrollToPanel('bulkParsedRow');
}

// =====================================================
// BULK SEARCH — Phase 3: Spreadsheet Preview & Column Mapping
// =====================================================

function bulkEscapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function bulkConvertSpacesToHash(value) {
    if (!value || typeof value !== 'string') return value;
    return value.trim().replace(/\s+/g, '#');
}

function bulkShowToast(id, message, containerId) {
    const existing = document.getElementById(id);
    if (existing) existing.remove();
    const container = document.getElementById(containerId);
    if (!container) return;
    const toast = document.createElement('div');
    toast.id = id;
    toast.className = 'selection-flash';
    toast.textContent = message;
    toast.style.cssText = 'position:absolute;top:8px;left:50%;transform:translateX(-50%);z-index:100;padding:4px 12px;border-radius:4px;font-size:12px;font-weight:600;background:var(--color-accent);color:white;pointer-events:none;';
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 1500);
}

function bulkShowHideWarning(message) {
    bulkShowToast('bulkHideWarning', message, 'bulkPreviewSectionContainer');
}

// =====================================================
// BULK SEARCH — Phase 3c: Preview Zoom, Column Mapping & Render
// =====================================================

function bulkUpdatePreviewZoom() {
    const wrapper = document.getElementById('bulkPreviewTableWrapper');
    const zoomLabel = document.getElementById('bulkZoomLevel');

    wrapper.style.transform = `scale(${bulkState.previewZoom / 100})`;
    wrapper.style.transformOrigin = 'top left';
    zoomLabel.textContent = `${bulkState.previewZoom}%`;
}

function bulkZoomPreview(direction) {
    const minZoom = 25;
    const maxZoom = 150;
    const step = 5;

    bulkState.previewZoom += direction * step;
    bulkState.previewZoom = Math.max(minZoom, Math.min(maxZoom, bulkState.previewZoom));
    bulkState.userManuallyZoomed = true;
    bulkUpdatePreviewZoom();
}

/**
 * Determine the last column (0-based) that contains actual data
 * across the visible row range (header through last row).
 */
function bulkGetLastDataRow(rows) {
    for (let r = rows.length - 1; r >= 0; r--) {
        const row = rows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
            if (row[c] !== undefined && row[c] !== null && String(row[c]).trim() !== '') {
                return r;
            }
        }
    }
    return 0;
}

function bulkGetLastDataColumn(rows, startRow, endRow) {
    let maxCol = 0;
    for (let r = startRow; r < endRow && r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        for (let c = row.length - 1; c >= 0; c--) {
            if (row[c] !== undefined && row[c] !== null && String(row[c]).trim() !== '') {
                if (c > maxCol) maxCol = c;
                break;
            }
        }
    }
    return maxCol;
}

function bulkAutoFitPreview() {
    if (bulkState.userHasResized) return; // User manually resized, don't auto-fit

    const wrapper = document.getElementById('bulkPreviewTableWrapper');
    const table = document.getElementById('bulkPreviewTable');
    const previewEl = document.getElementById('bulkSpreadsheetPreview');
    const scrollArea = document.getElementById('bulkPreviewContainer');

    if (!table || !wrapper || !previewEl || !scrollArea) return;

    // Disable CSS transition so getBoundingClientRect() measures final transform instantly
    wrapper.style.transition = 'none';

    // --- Step 1: Auto-fit zoom (width) if user hasn't manually zoomed ---
    if (!bulkState.userManuallyZoomed) {
        // Temporarily reset zoom to measure natural table width
        wrapper.style.transform = 'scale(1)';
        const naturalTableWidth = table.offsetWidth;
        const panelWidth = scrollArea.clientWidth - 10; // Small padding buffer

        if (naturalTableWidth > 0 && panelWidth > 0) {
            // Calculate zoom that fits table width to panel width
            const fitZoom = Math.floor((panelWidth / naturalTableWidth) * 100);

            if (fitZoom >= 55) {
                // Content fits at 55% or is narrower — zoom IN to fill width (cap at 100%)
                bulkState.previewZoom = Math.min(fitZoom, 100);
            } else {
                // Content is wider than panel at 55% — keep 55%, allow horizontal scroll
                bulkState.previewZoom = 55;
            }
        }

        bulkUpdatePreviewZoom();
    }

    // --- Step 2: Auto-fit height (priority) — snap to show all selected rows ---
    // getBoundingClientRect() forces a synchronous layout reflow, so the zoom
    // transform set in Step 1 is already applied when we measure here.

    // Measure the fixed chrome (header, toolbar, legend) — everything except the scroll area
    const previewRect = previewEl.getBoundingClientRect();
    const scrollRect = scrollArea.getBoundingClientRect();
    const chromeHeight = scrollRect.top - previewRect.top; // height of header + toolbar above scroll area

    const legend = previewEl.querySelector('.preview-legend');
    const legendHeight = legend ? legend.offsetHeight : 0;

    // Get scaled table height from the actual rendered bounding rect
    const tableRect = table.getBoundingClientRect();
    const scaledTableHeight = tableRect.height;

    // Target height: chrome above + scaled table + legend below + padding + scrollbar
    const BOTTOM_PADDING = 2;
    const SCROLLBAR_HEIGHT = 16; // Account for horizontal scrollbar
    const MIN_HEIGHT = 80;
    const targetHeight = chromeHeight + scaledTableHeight + legendHeight + BOTTOM_PADDING + SCROLLBAR_HEIGHT;
    const finalHeight = Math.max(MIN_HEIGHT, Math.ceil(targetHeight));

    const currentHeight = previewEl.offsetHeight;

    if (Math.abs(finalHeight - currentHeight) > 2) {
        previewEl.style.height = `${finalHeight}px`;
        console.log(`[BulkAutoFit] Snap-fit: content needs ${Math.ceil(targetHeight)}px at ${bulkState.previewZoom}% zoom -> ${finalHeight}px (was ${currentHeight}px)`);
    }

    // Force reflow to commit the instant (non-animated) transform, then restore CSS transition
    wrapper.offsetHeight;
    wrapper.style.transition = '';
}

function bulkResetColumnDropdowns() {
    document.getElementById('bulkColumnSelect').innerHTML = '<option value="">Select column...</option>';
    document.getElementById('bulkQtyColumnSelect').innerHTML = '<option value="">None</option>';
    document.getElementById('bulkResellerPriceColumnSelect').innerHTML = '<option value="">None</option>';
    document.getElementById('bulkVpnColumnSelect').innerHTML = '<option value="">None</option>';
    document.getElementById('bulkMsrpColumnSelect').innerHTML = '<option value="">None</option>';
}

function bulkAppendAutoSelectOption(selectEl, colLetter, headerText, selectId) {
    const option = document.createElement('option');
    option.value = colLetter;
    option.textContent = headerText;
    selectEl.appendChild(option);
}

function bulkUpdateColumnSelectionDropdown(headerRowIndex, rules, skipAutoSelect) {
    if (!bulkState.fileRows || bulkState.fileRows.length === 0) return;

    var headers = bulkState.fileRows[headerRowIndex] || [];

    // Compute last data column across the visible row range
    var lastRowInput = document.getElementById('bulkLastRowInput');
    var lastRowIdx = lastRowInput && lastRowInput.value ? parseInt(lastRowInput.value) - 1 : null;
    var totalRows = bulkState.fileRows.length;
    var endRow = lastRowIdx !== null ? Math.min(lastRowIdx + 1, totalRows) : totalRows;
    bulkState.lastDataColumn = bulkGetLastDataColumn(bulkState.fileRows, headerRowIndex, endRow);
    var maxDropdownCols = bulkState.lastDataColumn + 1;

    var columnSelect = document.getElementById('bulkColumnSelect');
    var qtyColumnSelect = document.getElementById('bulkQtyColumnSelect');
    var resellerPriceColumnSelect = document.getElementById('bulkResellerPriceColumnSelect');
    var vpnColumnSelect = document.getElementById('bulkVpnColumnSelect');
    var msrpColumnSelect = document.getElementById('bulkMsrpColumnSelect');

    // Use rules if provided (from Supabase), otherwise fall back to hardcoded keywords
    var mpnKeywords = (rules && rules.mpn) ? rules.mpn : ['item number', 'mpn', 'part number', 'mfg part', 'manufacturer part', 'mfr part', 'mfr.', 'sku'];
    var qtyKeywords = (rules && rules.qty) ? rules.qty : ['qty', 'quantity'];
    var priceKeywords = (rules && rules.price) ? rules.price : ['reseller price', 'reseller', 'dealer price', 'our price', 'unit price', 'unit cost', 'customer price', 'contract price', 'wholesale price', 'net price'];
    var vpnKeywords = (rules && rules.vpn) ? rules.vpn : ['vpn', 'vendor part', 'ingram part'];
    var msrpKeywords = (rules && rules.msrp) ? rules.msrp : ['msrp', 'list price', 'list', 'retail price', 'suggested retail'];

    // Reset all dropdowns
    bulkResetColumnDropdowns();

    // --- Pass 1: Populate all dropdown options (no auto-selection yet) ---
    // Also build a lookup of which columns match which field keywords
    var mpnCandidates = [];
    var qtyCandidates = [];
    var priceCandidates = [];
    var vpnCandidates = [];
    var msrpCandidates = [];

    headers.forEach(function(header, index) {
        // Skip columns beyond the last data column
        if (index >= maxDropdownCols) return;

        var colLetter = index < 26 ? String.fromCharCode(65 + index) : 'Col' + (index + 1);
        var headerText = colLetter + ': ' + (header || '(empty)');
        var headerLower = String(header).toLowerCase().trim();

        // Check if this is an "Ext"/"Extended" field (line-total columns, not unit values)
        var isExtField = /\bext\b\.?|extended/i.test(String(header));
        // Check if this is a "List Price" or "MSRP" field (not a reseller/cost price)
        var isListOrMsrp = /\blist\b|\bmsrp\b/i.test(String(header));

        // Add option to each dropdown (always — user can manually pick any column)
        var mpnOption = document.createElement('option');
        mpnOption.value = index;
        mpnOption.textContent = headerText;
        columnSelect.appendChild(mpnOption);

        var qtyOption = document.createElement('option');
        qtyOption.value = index;
        qtyOption.textContent = headerText;
        qtyColumnSelect.appendChild(qtyOption);

        var priceOption = document.createElement('option');
        priceOption.value = index;
        priceOption.textContent = headerText;
        resellerPriceColumnSelect.appendChild(priceOption);

        var vpnOption = document.createElement('option');
        vpnOption.value = index;
        vpnOption.textContent = headerText;
        vpnColumnSelect.appendChild(vpnOption);

        var msrpOption = document.createElement('option');
        msrpOption.value = index;
        msrpOption.textContent = headerText;
        msrpColumnSelect.appendChild(msrpOption);

        // Record candidate matches for auto-selection (respecting Ext exclusion)
        if (mpnKeywords.some(function(v) { return headerLower.indexOf(v) !== -1; })) {
            mpnCandidates.push(index);
        }
        if (qtyKeywords.some(function(v) { return headerLower.indexOf(v) !== -1; })) {
            qtyCandidates.push(index);
        }
        if (!isExtField && !isListOrMsrp && priceKeywords.some(function(v) { return headerLower.indexOf(v) !== -1; })) {
            priceCandidates.push(index);
        }
        if (vpnKeywords.some(function(v) { return headerLower.indexOf(v) !== -1; })) {
            vpnCandidates.push(index);
        }
        if (!isExtField && msrpKeywords.some(function(v) { return headerLower.indexOf(v) !== -1; })) {
            msrpCandidates.push(index);
        }
    });

    // --- Pass 2: Auto-select columns with priority ordering and no duplicates ---
    // Skip if caller only wants to populate options (e.g. initial file load)
    if (skipAutoSelect) return;

    // Priority: MPN → QTY → Price → VPN → MSRP
    // Each column index can only be claimed by one field
    var usedColumns = {};

    // Helper: pick the first unclaimed candidate for a dropdown
    function autoSelectFirst(selectEl, candidates, fieldName) {
        for (var i = 0; i < candidates.length; i++) {
            var colIdx = candidates[i];
            if (!usedColumns[colIdx]) {
                selectEl.value = String(colIdx);
                usedColumns[colIdx] = fieldName;
                console.log('[BulkAutoMap] Auto-selected column', colIdx, 'for', fieldName);
                return;
            }
        }
    }

    autoSelectFirst(columnSelect, mpnCandidates, 'mpn');
    autoSelectFirst(qtyColumnSelect, qtyCandidates, 'qty');
    autoSelectFirst(resellerPriceColumnSelect, priceCandidates, 'price');
    autoSelectFirst(vpnColumnSelect, vpnCandidates, 'vpn');
    autoSelectFirst(msrpColumnSelect, msrpCandidates, 'msrp');
}

function bulkRenderSpreadsheetPreview() {
    if (!bulkState.fileRows || bulkState.fileRows.length === 0) {
        bulkHideSpreadsheetPreview();
        return;
    }

    const previewEl = document.getElementById('bulkSpreadsheetPreview');
    const previewInfo = document.getElementById('bulkPreviewInfo');
    const thead = document.getElementById('bulkPreviewTableHead');
    const tbody = document.getElementById('bulkPreviewTableBody');
    const headerRowInput = document.getElementById('bulkHeaderRowInput');
    const lastRowInput = document.getElementById('bulkLastRowInput');

    const headerRowIdx = parseInt(headerRowInput.value) - 1 || 0;
    const lastRowIdx = lastRowInput.value
        ? parseInt(lastRowInput.value) - 1
        : null;
    const totalRows = bulkState.fileRows.length;

    // Determine the row range to display:
    // - Start from header row (include it for context)
    // - End at last row if specified, otherwise cap at last row with content (skip trailing empty rows)
    const startRow = headerRowIdx;
    let endRow;
    if (lastRowIdx !== null) {
        endRow = Math.min(lastRowIdx + 1, totalRows); // +1 to include last row
    } else {
        const lastDataRow = bulkGetLastDataRow(bulkState.fileRows);
        endRow = Math.min(lastDataRow + 1, totalRows); // +1 to include last data row
    }
    const displayRowCount = endRow - startRow;

    // Build column-to-class map for per-group highlighting
    const colClassMap = new Map();
    const mpnColVal = document.getElementById('bulkColumnSelect').value;
    const qtyColVal = document.getElementById('bulkQtyColumnSelect').value;
    const priceColVal = document.getElementById('bulkResellerPriceColumnSelect').value;
    const vpnColVal = document.getElementById('bulkVpnColumnSelect').value;
    const msrpColVal = document.getElementById('bulkMsrpColumnSelect').value;
    if (mpnColVal !== '') colClassMap.set(parseInt(mpnColVal), 'mapped-mpn');
    if (vpnColVal !== '') colClassMap.set(parseInt(vpnColVal), 'mapped-vpn');
    if (msrpColVal !== '') colClassMap.set(parseInt(msrpColVal), 'mapped-pricing');
    if (priceColVal !== '') colClassMap.set(parseInt(priceColVal), 'mapped-pricing');
    if (qtyColVal !== '') colClassMap.set(parseInt(qtyColVal), 'mapped-pricing');

    // Show the preview row container and add visible class
    const previewRow = document.getElementById('bulkPreviewRow');
    if (previewRow) previewRow.style.display = '';
    previewEl.classList.add('visible');

    // Count visible vs hidden for info display
    const hiddenRowCount = [...bulkState.hiddenRows].filter(r => r > headerRowIdx && (lastRowIdx === null || r < lastRowIdx)).length;
    const hiddenColCount = bulkState.hiddenColumns.size;
    let rangeInfo =
        lastRowIdx !== null
            ? `Rows ${startRow + 1}-${endRow} (${displayRowCount} rows)`
            : `Rows ${startRow + 1}-${endRow} of ${totalRows}`;
    if (hiddenRowCount > 0 || hiddenColCount > 0) {
        rangeInfo += ` | Hidden: ${hiddenRowCount > 0 ? hiddenRowCount + ' rows' : ''}${hiddenRowCount > 0 && hiddenColCount > 0 ? ', ' : ''}${hiddenColCount > 0 ? hiddenColCount + ' cols' : ''}`;
    }
    previewInfo.textContent = rangeInfo;

    // Determine column count — only columns that contain actual data (skip trailing empty columns)
    // Scan full data range (not just preview-visible rows) to find all data columns
    const dataEndRow = lastRowIdx !== null ? Math.min(lastRowIdx + 1, totalRows) : totalRows;
    bulkState.lastDataColumn = bulkGetLastDataColumn(bulkState.fileRows, startRow, dataEndRow);
    const maxCols = bulkState.lastDataColumn + 1; // +1 because lastDataColumn is 0-based

    // Build table header with hidden column support
    let theadHtml = '<tr><th class="row-header">#</th>';
    for (let c = 0; c < maxCols; c++) {
        const colLetter =
            c < 26
                ? String.fromCharCode(65 + c)
                : 'A' + String.fromCharCode(65 + c - 26);
        const mappedClass = colClassMap.has(c) ? ` ${colClassMap.get(c)}` : '';
        const isHidden = bulkState.hiddenColumns.has(c);

        if (isHidden) {
            // Hidden column - show narrow "..." indicator
            theadHtml += `<th class="hidden-col" onclick="bulkUnhideColumn(${c})" data-col="${c}" title="Click to unhide column ${colLetter}">...</th>`;
        } else {
            theadHtml += `<th class="${mappedClass}" onclick="bulkHandleColumnClick(${c})" data-col="${c}">${colLetter}</th>`;
        }
    }
    theadHtml += '</tr>';
    thead.innerHTML = theadHtml;

    // Build table body with hidden row support
    let tbodyHtml = '';
    let consecutiveHiddenIndices = [];

    for (let r = startRow; r < endRow; r++) {
        const row = bulkState.fileRows[r] || [];
        const isHeaderRow = r === headerRowIdx;
        const isLastRow = lastRowIdx !== null && r === lastRowIdx;
        const isHiddenRow = bulkState.hiddenRows.has(r) && !isHeaderRow && !isLastRow;

        // Handle hidden rows - collect consecutive hidden row indices
        if (isHiddenRow) {
            consecutiveHiddenIndices.push(r);
            continue;
        }

        // Output hidden rows indicator if we had consecutive hidden rows
        if (consecutiveHiddenIndices.length > 0) {
            tbodyHtml += bulkBuildHiddenRowsHtml(consecutiveHiddenIndices, maxCols + 1);
            consecutiveHiddenIndices = [];
        }

        let rowClass = '';
        if (isHeaderRow) rowClass = 'header-row';
        else if (isLastRow) rowClass = 'last-row';

        tbodyHtml += `<tr class="${rowClass}" onclick="bulkHandleRowClick(${r})" data-row="${r}">`;
        tbodyHtml += `<td class="row-header">${r + 1}</td>`;

        for (let c = 0; c < maxCols; c++) {
            const cellValue = row[c] !== undefined ? String(row[c]) : '';
            const isHiddenCol = bulkState.hiddenColumns.has(c);

            if (isHiddenCol) {
                // Hidden column cell
                tbodyHtml += `<td class="hidden-col" onclick="event.stopPropagation(); bulkUnhideColumn(${c})" title="Click to unhide">...</td>`;
            } else {
                const displayValue =
                    cellValue.length > 20
                        ? cellValue.substring(0, 20) + '...'
                        : cellValue;
                const mappedClass = colClassMap.has(c) ? ` ${colClassMap.get(c)}` : '';
                tbodyHtml += `<td class="${mappedClass}" title="${bulkEscapeHtml(cellValue)}">${bulkEscapeHtml(displayValue)}</td>`;
            }
        }
        tbodyHtml += '</tr>';
    }

    // Handle trailing hidden rows
    if (consecutiveHiddenIndices.length > 0) {
        tbodyHtml += bulkBuildHiddenRowsHtml(consecutiveHiddenIndices, maxCols + 1);
    }

    tbody.innerHTML = tbodyHtml;

    // Scroll preview to top on render
    const previewContainer = document.getElementById('bulkPreviewContainer');
    if (previewContainer) {
        previewContainer.scrollTop = 0;
        previewContainer.scrollLeft = 0;
    }

    // Update zoom display
    bulkUpdatePreviewZoom();
    // Auto-fit zoom (width) and snap-fit height after render
    bulkAutoFitPreview();
}

function bulkHideSpreadsheetPreview() {
    const previewRow = document.getElementById('bulkPreviewRow');
    if (previewRow) previewRow.style.display = 'none';
    document.getElementById('bulkSpreadsheetPreview').classList.remove('visible');
    // Clear selection mode and reset hide state when hiding preview
    bulkClearSelectionMode();
    bulkState.hiddenColumns.clear();
    bulkState.hiddenRows.clear();
    bulkState.userHasResized = false;
}

function bulkBuildHiddenRowsHtml(indices, colSpan) {
    const count = indices.length;
    const indicesJson = JSON.stringify(indices);
    return `<tr class="hidden-rows-indicator"><td colspan="${colSpan}" data-indices='${indicesJson}' onclick="bulkToggleHiddenRowsDropdown(event, JSON.parse(this.dataset.indices))">... ${count} hidden row${count > 1 ? 's' : ''} (click to manage) ...</td></tr>`;
}

// =====================================================
// BULK SEARCH — Phase 3c: Selection Mode, Toolbar, Mapping & Parsed Preview
// =====================================================

function bulkToggleSelectionMode(mode) {
    const previewTable = document.getElementById('bulkPreviewTable');
    const allBtns = document.getElementById('bulkSelectionToolbar').querySelectorAll('.sel-btn');

    // If clicking the same mode, deactivate
    if (bulkState.selectionMode === mode) {
        bulkState.selectionMode = null;
        allBtns.forEach(btn => btn.classList.remove('active'));
        previewTable.classList.remove('mode-row', 'mode-col', 'mode-hide-col', 'mode-hide-row');
        return;
    }

    // Activate new mode
    bulkState.selectionMode = mode;
    allBtns.forEach(btn => btn.classList.remove('active'));
    document.getElementById('bulkSelectionToolbar').querySelector(`[data-mode="${mode}"]`).classList.add('active');

    // Set table class for hover styling
    previewTable.classList.remove('mode-row', 'mode-col', 'mode-hide-col', 'mode-hide-row');
    if (mode === 'headerRow' || mode === 'lastRow') {
        previewTable.classList.add('mode-row');
    } else if (mode === 'hideCol') {
        previewTable.classList.add('mode-hide-col');
    } else if (mode === 'hideRow') {
        previewTable.classList.add('mode-hide-row');
    } else {
        previewTable.classList.add('mode-col');
    }
}

function bulkClearSelectionMode() {
    bulkState.selectionMode = null;
    const previewTable = document.getElementById('bulkPreviewTable');
    const allBtns = document.getElementById('bulkSelectionToolbar').querySelectorAll('.sel-btn');

    allBtns.forEach(btn => btn.classList.remove('active'));
    previewTable.classList.remove('mode-row', 'mode-col', 'mode-hide-col', 'mode-hide-row');
}

function bulkHandleRowClick(rowIndex) {
    if (!bulkState.selectionMode) return;

    // Handle hide row mode
    if (bulkState.selectionMode === 'hideRow') {
        bulkHandleHideRowClick(rowIndex);
        return;
    }

    if (bulkState.selectionMode !== 'headerRow' && bulkState.selectionMode !== 'lastRow') return;

    const currentMode = bulkState.selectionMode; // Save before clearing
    const input = currentMode === 'headerRow'
        ? document.getElementById('bulkHeaderRowInput')
        : document.getElementById('bulkLastRowInput');

    // Set the value (1-based for display)
    input.value = rowIndex + 1;

    // Flash the input to show it changed
    input.classList.add('selection-flash');
    setTimeout(() => input.classList.remove('selection-flash'), 500);

    // Update the dropdown options if header row changed
    if (currentMode === 'headerRow') {
        bulkUpdateColumnSelectionDropdown(rowIndex);
    }

    // Reset userHasResized when header/last row changes
    bulkState.userHasResized = false;

    // Clear mode and re-render
    bulkClearSelectionMode();
    bulkOnRowInputChange();

    console.log(`[BulkSearch] Set ${currentMode} to row ${rowIndex + 1}`);
}

function bulkHandleHideRowClick(rowIndex) {
    const headerRowInput = document.getElementById('bulkHeaderRowInput');
    const lastRowInput = document.getElementById('bulkLastRowInput');
    const headerRowIdx = parseInt(headerRowInput.value) - 1 || 0;
    const lastRowIdx = lastRowInput.value ? parseInt(lastRowInput.value) - 1 : null;

    // Can't hide header or last row
    if (rowIndex === headerRowIdx) {
        bulkShowHideWarning('Cannot hide the header row');
        return;
    }
    if (lastRowIdx !== null && rowIndex === lastRowIdx) {
        bulkShowHideWarning('Cannot hide the last row');
        return;
    }
    // Can only hide rows between header and last
    if (rowIndex < headerRowIdx || (lastRowIdx !== null && rowIndex > lastRowIdx)) {
        bulkShowHideWarning('Can only hide rows between header and last row');
        return;
    }

    // Toggle hidden state
    if (bulkState.hiddenRows.has(rowIndex)) {
        bulkState.hiddenRows.delete(rowIndex);
        console.log(`[BulkSearch] Unhid row ${rowIndex + 1}`);
    } else {
        bulkState.hiddenRows.add(rowIndex);
        console.log(`[BulkSearch] Hid row ${rowIndex + 1}`);
    }

    // Re-render (don't clear mode - allow multiple selections)
    bulkRenderSpreadsheetPreview();
}

function bulkHandleColumnClick(colIndex) {
    if (!bulkState.selectionMode) return;

    // Handle hide column mode
    if (bulkState.selectionMode === 'hideCol') {
        bulkHandleHideColumnClick(colIndex);
        return;
    }

    if (
        bulkState.selectionMode !== 'mpnCol' &&
        bulkState.selectionMode !== 'qtyCol' &&
        bulkState.selectionMode !== 'priceCol' &&
        bulkState.selectionMode !== 'vpnCol' &&
        bulkState.selectionMode !== 'msrpCol'
    ) return;

    const currentMode = bulkState.selectionMode; // Save before clearing
    const MODE_TO_SELECT = {
        mpnCol: 'bulkColumnSelect',
        qtyCol: 'bulkQtyColumnSelect',
        vpnCol: 'bulkVpnColumnSelect',
        msrpCol: 'bulkMsrpColumnSelect',
        priceCol: 'bulkResellerPriceColumnSelect'
    };
    const selectId = MODE_TO_SELECT[currentMode];
    const select = document.getElementById(selectId);

    // Toggle: if clicking the already-mapped column, unset it
    if (select.value === String(colIndex)) {
        select.value = '';
        select.classList.add('selection-flash');
        setTimeout(() => select.classList.remove('selection-flash'), 500);
        bulkClearSelectionMode();
        bulkOnMappingChange();
        console.log(`[BulkSearch] Unset ${currentMode} (was column ${colIndex})`);
        return;
    }

    // Guard: VPN column cannot be the same as MPN column
    if (currentMode === 'vpnCol') {
        const mpnCol = document.getElementById('bulkColumnSelect').value;
        if (mpnCol !== '' && String(colIndex) === String(mpnCol)) {
            bulkShowHideWarning('VPN column cannot be the same as MPN column');
            bulkClearSelectionMode();
            return;
        }
    }

    // Guard: MPN column cannot be the same as VPN column
    if (currentMode === 'mpnCol') {
        const vpnCol = document.getElementById('bulkVpnColumnSelect').value;
        if (vpnCol !== '' && String(colIndex) === String(vpnCol)) {
            bulkShowHideWarning('MPN column cannot be the same as VPN column');
            bulkClearSelectionMode();
            return;
        }
    }

    // Set the value
    select.value = colIndex;

    // Flash the select to show it changed
    select.classList.add('selection-flash');
    setTimeout(() => select.classList.remove('selection-flash'), 500);

    // Clear mode and re-render
    bulkClearSelectionMode();
    bulkOnMappingChange();

    console.log(`[BulkSearch] Set ${currentMode} to column ${colIndex}`);
}

function bulkHandleHideColumnClick(colIndex) {
    // Check if this is a mapped column
    const guards = [
        ['bulkColumnSelect', 'MPN'],
        ['bulkQtyColumnSelect', 'QTY'],
        ['bulkResellerPriceColumnSelect', 'Price'],
        ['bulkVpnColumnSelect', 'VPN'],
        ['bulkMsrpColumnSelect', 'MSRP']
    ];
    for (const [id, label] of guards) {
        const val = document.getElementById(id).value;
        if (val !== '' && parseInt(val) === colIndex) {
            bulkShowHideWarning('Cannot hide ' + label + ' column');
            return;
        }
    }

    // Toggle hidden state
    if (bulkState.hiddenColumns.has(colIndex)) {
        bulkState.hiddenColumns.delete(colIndex);
        console.log(`[BulkSearch] Unhid column ${colIndex}`);
    } else {
        bulkState.hiddenColumns.add(colIndex);
        console.log(`[BulkSearch] Hid column ${colIndex}`);
    }

    // Re-render (don't clear mode - allow multiple selections)
    bulkRenderSpreadsheetPreview();
}

function bulkUnhideColumn(colIndex) {
    bulkState.hiddenColumns.delete(colIndex);
    bulkRenderSpreadsheetPreview();
    console.log(`[BulkSearch] Unhid column ${colIndex}`);
}

function bulkUnhideRow(rowIndex) {
    bulkState.hiddenRows.delete(rowIndex);
    bulkCloseHiddenRowsDropdown();
    bulkRenderSpreadsheetPreview();
    console.log(`[BulkSearch] Unhid row ${rowIndex + 1}`);
}

function bulkUnhideAllRows() {
    bulkState.hiddenRows.clear();
    bulkCloseHiddenRowsDropdown();
    bulkRenderSpreadsheetPreview();
    console.log('[BulkSearch] Unhid all rows');
}

function bulkToggleHiddenRowsDropdown(event, hiddenRowIndices) {
    event.stopPropagation();

    // Close existing dropdown if any
    if (bulkState.activeHiddenRowsDropdown) {
        bulkCloseHiddenRowsDropdown();
        return;
    }

    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.className = 'hidden-rows-dropdown';

    // Add individual row items
    hiddenRowIndices.forEach(rowIdx => {
        const item = document.createElement('div');
        item.className = 'hidden-row-item';
        item.innerHTML = `
            <span class="row-label">Row ${rowIdx + 1}</span>
            <button class="unhide-row-btn" onclick="event.stopPropagation(); bulkUnhideRow(${rowIdx})">Unhide</button>
        `;
        dropdown.appendChild(item);
    });

    // Add "Unhide All" button
    const unhideAllBtn = document.createElement('button');
    unhideAllBtn.className = 'unhide-all-btn';
    unhideAllBtn.textContent = `Unhide All (${hiddenRowIndices.length})`;
    unhideAllBtn.onclick = (e) => {
        e.stopPropagation();
        bulkUnhideAllRows();
    };
    dropdown.appendChild(unhideAllBtn);

    // Position dropdown relative to clicked cell
    const cell = event.currentTarget;
    const rect = cell.getBoundingClientRect();

    dropdown.style.position = 'fixed';
    dropdown.style.top = `${rect.bottom + 4}px`;
    dropdown.style.left = `${rect.left + rect.width / 2}px`;
    dropdown.style.transform = 'translateX(-50%)';

    document.body.appendChild(dropdown);
    bulkState.activeHiddenRowsDropdown = dropdown;

    // Close on outside click (delayed to avoid immediate close)
    setTimeout(() => {
        document.addEventListener('click', bulkHandleDropdownOutsideClick);
    }, 0);
}

function bulkCloseHiddenRowsDropdown() {
    if (bulkState.activeHiddenRowsDropdown) {
        bulkState.activeHiddenRowsDropdown.remove();
        bulkState.activeHiddenRowsDropdown = null;
    }
    document.removeEventListener('click', bulkHandleDropdownOutsideClick);
}

function bulkHandleDropdownOutsideClick(event) {
    if (bulkState.activeHiddenRowsDropdown && !bulkState.activeHiddenRowsDropdown.contains(event.target)) {
        bulkCloseHiddenRowsDropdown();
    }
}

// =====================================================
// BULK SEARCH — Phase 3c: Mapping Change Handlers
// =====================================================

function bulkOnRowInputChange() {
    // Reset userHasResized when header/last row changes (re-enable auto-fit)
    bulkState.userHasResized = false;

    if (bulkState.fileRows && bulkState.fileRows.length > 0) {
        bulkRenderSpreadsheetPreview();
    }
}

function bulkOnMappingChange() {
    if (bulkState.fileRows && bulkState.fileRows.length > 0) {
        bulkRenderSpreadsheetPreview();
    }
    bulkUpdateParsedPreview();
}

function bulkResetColumnMappings() {
    // Reset header/last row inputs
    document.getElementById('bulkHeaderRowInput').value = '1';
    document.getElementById('bulkLastRowInput').value = '';

    // Reset all column dropdowns
    bulkResetColumnDropdowns();

    // Clear any active selection mode
    bulkClearSelectionMode();

    // Clear hidden columns and rows
    bulkState.hiddenColumns.clear();
    bulkState.hiddenRows.clear();

    // Re-run column auto-detect if file is loaded
    if (bulkState.fileRows && bulkState.fileRows.length > 0) {
        bulkUpdateColumnSelectionDropdown(0);
    }

    // Re-render preview with fresh mappings
    bulkRenderSpreadsheetPreview();
    bulkUpdateParsedPreview();

    console.log('[BulkSearch] Reset all column mappings');
}

// =====================================================
// BULK SEARCH — Phase 3c: Apply Column Selection & Parsed Preview
// =====================================================

function bulkApplyColumnSelection() {
    const columnSelect = document.getElementById('bulkColumnSelect');
    const headerRowInput = document.getElementById('bulkHeaderRowInput');
    const lastRowInput = document.getElementById('bulkLastRowInput');
    const qtyColumnSelect = document.getElementById('bulkQtyColumnSelect');
    const resellerPriceColumnSelect = document.getElementById('bulkResellerPriceColumnSelect');
    const vpnColumnSelect = document.getElementById('bulkVpnColumnSelect');
    const msrpColumnSelect = document.getElementById('bulkMsrpColumnSelect');

    const selectedColumn = parseInt(columnSelect.value);
    const headerRow = parseInt(headerRowInput.value) - 1 || 0;
    const lastRow = lastRowInput.value ? parseInt(lastRowInput.value) : null;
    const qtyColumn = qtyColumnSelect.value !== '' ? parseInt(qtyColumnSelect.value) : null;
    const resellerPriceColumn = resellerPriceColumnSelect.value !== '' ? parseInt(resellerPriceColumnSelect.value) : null;
    const vpnColumn = vpnColumnSelect.value !== '' ? parseInt(vpnColumnSelect.value) : null;
    const msrpColumn = msrpColumnSelect.value !== '' ? parseInt(msrpColumnSelect.value) : null;

    if (isNaN(selectedColumn) || !bulkState.fileRows) {
        alert('Please select an MPN column');
        return;
    }

    // Determine end row (exclusive)
    const startRow = headerRow + 1;
    const endRow = lastRow
        ? Math.min(lastRow, bulkState.fileRows.length)
        : bulkState.fileRows.length;

    // Extract data from selected columns, starting after header row up to last row
    const dataRows = bulkState.fileRows.slice(startRow, endRow);

    // Build parsed data with MPN, qty, and reseller price
    const parsedData = [];
    const seenMpns = new Set();

    dataRows.forEach((row, rowIndex) => {
        // Skip hidden rows
        const absoluteRowIndex = startRow + rowIndex;
        if (bulkState.hiddenRows.has(absoluteRowIndex)) return;

        const mpnRaw = row[selectedColumn];
        if (!mpnRaw || !String(mpnRaw).trim()) return; // Skip empty rows

        const mpn = String(bulkConvertSpacesToHash(String(mpnRaw))).toUpperCase();
        if (seenMpns.has(mpn)) return; // Skip duplicates
        seenMpns.add(mpn);

        // Get QTY from column if selected
        let qty = 1;
        if (qtyColumn !== null && row[qtyColumn]) {
            const parsedQty = parseInt(row[qtyColumn]);
            if (!isNaN(parsedQty) && parsedQty >= 1 && parsedQty <= 9999) {
                qty = parsedQty;
            }
        }

        // Get Reseller Price from column if selected
        let resellerPrice = null;
        if (resellerPriceColumn !== null && row[resellerPriceColumn]) {
            const priceStr = String(row[resellerPriceColumn]).replace(/[$,]/g, '');
            const parsedPrice = parseFloat(priceStr);
            if (!isNaN(parsedPrice) && parsedPrice >= 0) {
                resellerPrice = parsedPrice;
            }
        }

        // Extract VPN from VPN column if set
        let vpn = null;
        if (vpnColumn !== null && row[vpnColumn] !== undefined && row[vpnColumn] !== null) {
            const rawVpn = String(row[vpnColumn]).trim();
            if (rawVpn) vpn = rawVpn;
        }

        // Extract MSRP from MSRP column if set
        let msrp = null;
        if (msrpColumn !== null && row[msrpColumn]) {
            const msrpStr = String(row[msrpColumn]).replace(/[$,]/g, '');
            const parsedMsrp = parseFloat(msrpStr);
            if (!isNaN(parsedMsrp) && parsedMsrp >= 0) {
                msrp = parsedMsrp;
            }
        }

        parsedData.push({ mpn, vpn, qty, resellerPrice, msrp });
    });

    if (parsedData.length === 0) {
        alert('No MPNs found in selected column');
        return;
    }

    // Store parsed data with metadata for product generation
    bulkState.parsedFileData = parsedData;

    // Update parsedSkus for display (just the MPNs) — merge with paste SKUs, deduplicated
    const newMpns = parsedData.map(d => d.mpn);
    bulkState.parsedSkus = [...new Set([...bulkState.parsedSkus, ...newMpns])];
    bulkUpdateParsedPreview();
    updateFooterStats();

    // Show parsed row
    const parsedRow = document.getElementById('bulkParsedRow');
    if (parsedRow) parsedRow.style.display = '';
    scrollToPanel('bulkParsedRow');

    // Count hidden rows in range for logging
    const hiddenInRange = [...bulkState.hiddenRows].filter(r => r >= startRow && r < endRow).length;
    console.log(
        `[BulkSearch] Parsed ${parsedData.length} MPNs from rows ${startRow + 1} to ${endRow}` +
        (hiddenInRange > 0 ? ` (${hiddenInRange} hidden rows excluded)` : '') +
        (vpnColumn !== null ? ', with VPN column' : '') +
        (msrpColumn !== null ? ', with MSRP column' : '') +
        (qtyColumn !== null ? ', with QTY column' : '') +
        (resellerPriceColumn !== null ? ', with Reseller Price column' : '')
    );

    // Show action bar when we have SKUs
    const actionBar = document.getElementById('bulkActionBar');
    if (actionBar && bulkState.parsedSkus.length > 0) actionBar.style.display = '';

    // Fire-and-forget: save column mappings if new keywords detected
    bulkSaveColumnMappingsIfNeeded();
}

function bulkSaveColumnMappingsIfNeeded() {
    var distNames = { ingram: 'Ingram Micro', tdsynnex: 'TD Synnex', adi: 'ADI Global' };
    var distributor = state.currentDistributor;
    if (!distributor) {
        console.log('[BulkSave] No distributor selected, skipping save');
        return;
    }
    var displayName = distNames[distributor] || distributor;

    // 1. Gather current selections — read header labels from the header row
    var headerRowInput = document.getElementById('bulkHeaderRowInput');
    var headerRow = parseInt(headerRowInput.value) - 1 || 0;
    if (!bulkState.fileRows || !bulkState.fileRows[headerRow]) {
        console.log('[BulkSave] No header row data available, skipping save');
        return;
    }
    var headerRowData = bulkState.fileRows[headerRow];

    // Map dropdown IDs to field names
    var dropdownMap = {
        mpn: 'bulkColumnSelect',
        qty: 'bulkQtyColumnSelect',
        price: 'bulkResellerPriceColumnSelect',
        vpn: 'bulkVpnColumnSelect',
        msrp: 'bulkMsrpColumnSelect'
    };

    // Collect user-selected header labels for each field
    var userSelections = {};
    var hasAnySelection = false;
    var fields = ['mpn', 'qty', 'price', 'vpn', 'msrp'];
    for (var i = 0; i < fields.length; i++) {
        var field = fields[i];
        var el = document.getElementById(dropdownMap[field]);
        if (!el || el.value === '' || el.value === 'none') continue;
        var colIdx = parseInt(el.value);
        if (isNaN(colIdx)) continue;
        var label = headerRowData[colIdx];
        if (!label || !String(label).trim()) continue;
        userSelections[field] = String(label).trim().toLowerCase();
        hasAnySelection = true;
    }

    if (!hasAnySelection) {
        console.log('[BulkSave] No column selections to save');
        return;
    }

    console.log('[BulkSave] User selections:', userSelections);

    // Map internal distributor ID to database column value
    var dbDistributorName = distributor === 'adi' ? 'adiglobal' : distributor;

    // 2. Fetch current Supabase row for this distributor
    var fetchUrl = SUPABASE_URL + '/rest/v1/bulk_column_mapping_rules?distributor=eq.' +
        encodeURIComponent(dbDistributorName) + '&select=distributor,mpn,qty,price,vpn,msrp';

    fetch(fetchUrl, {
        headers: {
            'apikey': SUPABASE_ANON_KEY,
            'Authorization': 'Bearer ' + SUPABASE_ANON_KEY
        }
    })
    .then(function(res) { return res.json(); })
    .then(function(rows) {
        var existingRow = (rows && rows.length > 0) ? rows[0] : null;
        if (!existingRow) {
            console.log('[BulkSave] No existing row for distributor:', distributor);
            return;
        }

        // 3. Compare — check if ALL mapped labels already exist
        var newKeywords = {}; // field -> new keyword to add
        var allExist = true;
        var selectedFields = Object.keys(userSelections);
        for (var i = 0; i < selectedFields.length; i++) {
            var field = selectedFields[i];
            var userLabel = userSelections[field];
            var existingStr = existingRow[field] || '';
            var existingKeywords = existingStr.split(',').map(function(k) { return k.trim().toLowerCase(); }).filter(Boolean);
            if (existingKeywords.indexOf(userLabel) === -1) {
                allExist = false;
                newKeywords[field] = userLabel;
            }
        }

        if (allExist) {
            console.log('[BulkSave] All mappings already saved, no update needed');
            return;
        }

        console.log('[BulkSave] New keywords to save:', newKeywords);

        // 4. Prompt user
        if (!confirm('Save mappings for future ' + displayName + ' quotes?')) {
            console.log('[BulkSave] User declined to save mappings');
            return;
        }

        // 5. Build PATCH body — only fields that changed
        var patchBody = { updated_at: new Date().toISOString() };
        var changedFields = Object.keys(newKeywords);
        for (var i = 0; i < changedFields.length; i++) {
            var field = changedFields[i];
            var existingStr = existingRow[field] || '';
            var existingKeywords = existingStr.split(',').map(function(k) { return k.trim(); }).filter(Boolean);
            existingKeywords.push(newKeywords[field]);
            patchBody[field] = existingKeywords.join(', ');
        }

        console.log('[BulkSave] Saving mappings:', patchBody);

        var patchUrl = SUPABASE_URL + '/rest/v1/bulk_column_mapping_rules?distributor=eq.' +
            encodeURIComponent(dbDistributorName);

        return fetch(patchUrl, {
            method: 'PATCH',
            headers: {
                'apikey': SUPABASE_ANON_KEY,
                'Authorization': 'Bearer ' + SUPABASE_ANON_KEY,
                'Content-Type': 'application/json',
                'Prefer': 'return=minimal'
            },
            body: JSON.stringify(patchBody)
        });
    })
    .then(function(res) {
        if (res && !res.ok) {
            console.error('[BulkSave] Supabase PATCH failed with status:', res.status);
            return;
        }
        if (res) {
            console.log('[BulkSave] Mappings saved successfully');
            showStatus('Mappings successfully saved. Auto Map will now use saved settings for future ' + displayName + ' quotes.', 'success');
        }
    })
    .catch(function(err) {
        console.error('[BulkSave] Error saving mappings:', err);
    });
}

function bulkUpdateParsedPreview() {
    const editableEl = document.getElementById('bulkParsedSkusEditable');
    const countEl = document.getElementById('bulkParsedCount');

    if (!editableEl || !countEl) return;

    countEl.textContent = `${bulkState.parsedSkus.length} SKUs`;

    // Display as comma-separated values in editable textarea
    if (bulkState.parsedSkus.length === 0) {
        editableEl.value = '';
        editableEl.placeholder = 'No MPNs parsed yet - values will appear here comma-separated';
    } else {
        editableEl.value = bulkState.parsedSkus.join(', ');
    }

    // Show parsed row FIRST so element is visible for scrollHeight measurement
    if (bulkState.parsedSkus.length > 0) {
        const parsedRow = document.getElementById('bulkParsedRow');
        if (parsedRow) parsedRow.style.display = '';
    }

    const actionBar = document.getElementById('bulkActionBar');
    if (actionBar) actionBar.style.display = bulkState.parsedSkus.length > 0 ? '' : 'none';

    // Auto-expand textarea AFTER parsedRow is visible — use rAF to ensure layout is complete
    requestAnimationFrame(function() {
        var el = document.getElementById('bulkParsedSkusEditable');
        if (!el) return;
        el.style.height = 'auto';
        el.style.height = Math.max(48, el.scrollHeight + 4) + 'px';
    });
}

function bulkUpdateParsedFromEdit() {
    const editableEl = document.getElementById('bulkParsedSkusEditable');
    if (!editableEl) return;

    const text = editableEl.value.trim();

    if (!text) {
        bulkState.parsedSkus = [];
        bulkState.parsedFileData = null; // Clear file data when manually editing
    } else {
        // Parse comma-separated values, dedupe and clean
        const splitPattern = /[,\s\t\n]+/;
        bulkState.parsedSkus = [
            ...new Set(
                text
                    .split(splitPattern)
                    .map(s => s.trim().toUpperCase())
                    .filter(s => s.length > 0)
            )
        ];
        // Manual edits break the link to file data (qty/price will default)
        // Only clear if user actually changed the content
        if (bulkState.parsedFileData) {
            const fileDataMpns = new Set(bulkState.parsedFileData.map(d => d.mpn));
            // Check if sets are different
            if (
                bulkState.parsedSkus.length !== bulkState.parsedFileData.length ||
                !bulkState.parsedSkus.every(mpn => fileDataMpns.has(mpn))
            ) {
                bulkState.parsedFileData = null;
                console.log('[BulkSearch] Manual edit detected, file data cleared - qty/price will default');
            }
        }
    }

    // Update count only (don't rewrite the textarea while editing)
    const countEl = document.getElementById('bulkParsedCount');
    if (countEl) countEl.textContent = `${bulkState.parsedSkus.length} SKUs`;
}

// =====================================================
// BULK SEARCH — Phase 4: RPC Search & Product Loading
// =====================================================

function bulkGetBatchSize() {
    const input = document.getElementById('bulkBatchSizeInput');
    return input ? Math.min(Math.max(parseInt(input.value) || 50, 1), 500) : 50;
}

async function bulkFetchRpcBatch(fnName, bodyPayload) {
    const resp = await fetch(`${SUPABASE_URL}/rest/v1/rpc/${fnName}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'apikey': SUPABASE_ANON_KEY,
            'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
        },
        body: JSON.stringify(bodyPayload)
    });
    if (!resp.ok) {
        const errText = await resp.text();
        throw new Error(`RPC ${fnName} failed (${resp.status}): ${errText}`);
    }
    const data = await resp.json();
    return Array.isArray(data) ? data : [];
}

function bulkFetchMpnBatch(mpns) {
    const rpcMap = {
        ingram: 'bulk_mpn_lookup_ingram',
        tdsynnex: 'bulk_mpn_lookup_tdsynnex',
        adi: 'bulk_mpn_lookup_adi',
    };
    const fnName = rpcMap[state.currentDistributor];
    if (!fnName) throw new Error(`Unknown distributor: ${state.currentDistributor}`);
    return bulkFetchRpcBatch(fnName, { mpns });
}

function bulkFetchVpnBatch(vpns) {
    return bulkFetchRpcBatch('bulk_ingram_vpn_lookup', { vpns });
}

function bulkMapRpcRowToProduct(row, distributor) {
    switch (distributor) {
        case 'ingram':
            return {
                mpn: row.vendor_part_number || '',
                vpn: row.ingram_part_number || '',
                manufacturer: row.manufacturer || row.vendor_name || '',
                description: row.description_line_1 || '',
                msrp: parseFloat(row.retail_price) || 0,
                resellerPrice: parseFloat(row.customer_price) || 0,
                category: row.category || '',
                _source: 'ingram',
            };
        case 'tdsynnex':
            return {
                mpn: row.manufacturer_part_number || '',
                vpn: row.synnex_sku || row.td_synnex_sku || '',
                manufacturer: row.manufacturer_name || '',
                description: row.part_description || row.long_description_1 || '',
                msrp: parseFloat(row.msrp) || 0,
                resellerPrice: parseFloat(row.contract_price) || 0,
                category: row.category_description || '',
                _source: 'tdsynnex',
            };
        case 'adi':
            return {
                mpn: row.product_code_mpn || row.vendor_part_code || '',
                vpn: row.item || '',
                manufacturer: row.manufacturer || '',
                description: row.item_desc || '',
                msrp: parseFloat(row.msrp) || 0,
                resellerPrice: parseFloat(row.current_price) || 0,
                category: row.category_1 || '',
                _source: 'adi',
            };
        default:
            return { mpn: '', manufacturer: '', description: '', msrp: 0, resellerPrice: 0, _source: 'ingram' };
    }
}

async function bulkLoadProducts() {
    if (bulkState.isLoading) return;
    if (bulkState.parsedSkus.length === 0) {
        alert('No MPNs to search. Please upload a file or paste MPNs first.');
        return;
    }

    bulkState.isLoading = true;

    // Deduplicate MPNs
    const uniqueMpns = [...new Set(bulkState.parsedSkus.map(m => m.trim().toUpperCase()))].filter(Boolean);

    // Determine if VPN lookup is available (ingram + file data + VPN values present)
    const useVpnLookup = (
        state.currentDistributor === 'ingram' &&
        bulkState.parsedFileData &&
        bulkState.parsedFileData.length > 0 &&
        bulkState.parsedFileData.some(d => d.vpn)
    );

    // Disable Load button
    const loadBtn = document.getElementById('bulkLoadProductsBtn');
    const originalBtnHTML = loadBtn ? loadBtn.innerHTML : '';
    if (loadBtn) {
        loadBtn.disabled = true;
        loadBtn.innerHTML = 'Loading\u2026';
    }

    // Show progress bar
    const progressEl = document.getElementById('bulkProgress');
    const progressFill = document.getElementById('bulkProgressFill');
    const progressText = document.getElementById('bulkProgressText');
    if (progressEl) progressEl.classList.add('visible');
    if (progressFill) progressFill.style.width = '0%';
    if (progressText) progressText.textContent = 'Starting...';

    // Clear previous results
    bulkState.products = [];
    bulkState.unmatchedMpns = [];
    bulkState.rawRpcRows.clear();
    bulkClearUnmatchedDisplay();

    try {
        const batchSize = bulkGetBatchSize();
        const allRows = [];
        const totalBatches = Math.ceil(uniqueMpns.length / batchSize);

        // Chunk MPNs per batch size (sequential to avoid overwhelming the API)
        for (let i = 0; i < uniqueMpns.length; i += batchSize) {
            const batchNum = Math.floor(i / batchSize) + 1;
            const chunk = uniqueMpns.slice(i, i + batchSize);

            // Update progress
            const pct = Math.round((batchNum / totalBatches) * 100);
            if (progressFill) progressFill.style.width = `${pct}%`;
            if (progressText) progressText.textContent = `Batch ${batchNum}/${totalBatches}`;

            let rows;

            if (useVpnLookup) {
                // Guard: VPN and MPN columns must not be the same
                const vpnColIdx = document.getElementById('bulkVpnColumnSelect')?.value;
                const mpnColIdx = document.getElementById('bulkColumnSelect')?.value;
                if (vpnColIdx && mpnColIdx && vpnColIdx === mpnColIdx) {
                    throw new Error('VPN and MPN columns are set to the same column. Please re-select your columns.');
                }

                // Build VPN array parallel to this MPN chunk
                const mpnToVpn = new Map(
                    bulkState.parsedFileData
                        .filter(d => d.vpn)
                        .map(d => [d.mpn.trim().toUpperCase(), d.vpn])
                );
                const vpnChunk = chunk
                    .map(mpn => mpnToVpn.get(mpn))
                    .filter(Boolean);

                if (vpnChunk.length > 0) {
                    rows = await bulkFetchVpnBatch(vpnChunk);
                    if (rows.length === 0) {
                        throw new Error(
                            'VPN lookup returned no results. Please verify:\n' +
                            '1. The VPN column is set to the Ingram Part Number column (not MPN)\n' +
                            '2. The uploaded file is an Ingram Micro distributor quote\n' +
                            '3. The products exist in the Ingram catalog\n\n' +
                            'First 3 VPN values sent: ' + vpnChunk.slice(0, 3).join(', ')
                        );
                    }
                } else {
                    rows = await bulkFetchMpnBatch(chunk);
                }
            } else {
                rows = await bulkFetchMpnBatch(chunk);
            }

            allRows.push(...rows);

            // Store raw RPC rows for product details lookup
            rows.forEach(row => {
                let mpnKey = '';
                if (state.currentDistributor === 'ingram') {
                    mpnKey = (row.vendor_part_number || '').toUpperCase();
                } else if (state.currentDistributor === 'adi') {
                    mpnKey = (row.product_code_mpn || '').toUpperCase();
                } else {
                    mpnKey = (row.manufacturer_part_number || '').toUpperCase();
                }
                if (mpnKey) bulkState.rawRpcRows.set(mpnKey, row);
            });
        }

        // Map RPC rows to product format
        bulkState.products = allRows.map(row => bulkMapRpcRowToProduct(row, state.currentDistributor));

        // Lazy verification — fire-and-forget, non-blocking
        if (state.currentDistributor === 'ingram') {
            verifyIngramManufacturers(bulkState.products);
        }

        // Merge in qty, resellerPrice, and msrp from file data if available
        if (bulkState.parsedFileData && bulkState.parsedFileData.length > 0) {
            const fileDataMap = new Map(bulkState.parsedFileData.map(d => [d.mpn.trim().toUpperCase(), d]));
            bulkState.products.forEach(p => {
                const fd = fileDataMap.get(p.mpn.toUpperCase());
                if (fd) {
                    if (fd.qty) p.qty = fd.qty;
                    if (fd.resellerPrice) p.resellerPrice = fd.resellerPrice;
                    if (fd.msrp !== null && fd.msrp !== undefined) {
                        p._dbMsrp = p.msrp;      // Save database MSRP
                        p._fileMsrp = fd.msrp;    // Save spreadsheet MSRP
                        p.msrp = fd.msrp;         // Overlay (will be corrected by comparison UI if needed)
                    }
                    if (fd.vpn) p._fileVpn = fd.vpn;
                }
            });
        }

        // Phase 7b: Detect MSRP mismatches
        bulkState.msrpMismatches = [];
        bulkState.msrpChoices = new Map();
        bulkState.products.forEach(p => {
            if (p._dbMsrp != null && p._dbMsrp !== 0 && p._fileMsrp != null && p._fileMsrp !== 0 && p._dbMsrp !== p._fileMsrp) {
                bulkState.msrpMismatches.push({
                    mpn: p.mpn,
                    description: p.description,
                    fileMsrp: p._fileMsrp,
                    dbMsrp: p._dbMsrp,
                    difference: p._dbMsrp - p._fileMsrp
                });
                // Smart default: pre-select the LOWER price (higher margin)
                bulkState.msrpChoices.set(p.mpn.toUpperCase(), p._dbMsrp <= p._fileMsrp ? 'current' : 'quote');
            }
        });

        // Restore original spreadsheet order (RPC results may arrive in any order)
        if (bulkState.parsedFileData && bulkState.parsedFileData.length > 0) {
            const orderMap = new Map();
            bulkState.parsedFileData.forEach((item, index) => {
                orderMap.set(item.mpn.trim().toUpperCase(), index);
            });
            bulkState.products.sort((a, b) => {
                const orderA = orderMap.get((a.mpn || '').toUpperCase());
                const orderB = orderMap.get((b.mpn || '').toUpperCase());
                return (orderA !== undefined ? orderA : 999999) - (orderB !== undefined ? orderB : 999999);
            });
        }

        // Identify unmatched MPNs
        const matchedMpns = new Set(bulkState.products.map(p => p.mpn.toUpperCase()));
        bulkState.unmatchedMpns = uniqueMpns.filter(m => !matchedMpns.has(m));

        // Show unmatched MPNs if any
        if (bulkState.unmatchedMpns.length > 0) {
            bulkShowUnmatchedDisplay(bulkState.unmatchedMpns);
        }

        console.log(`[BulkSearch] Loaded ${bulkState.products.length} products, ${bulkState.unmatchedMpns.length} unmatched`);
        updateFooterStats();

        // Reset pagination for new search results
        bulkState.resultsPage = 1;

        // Phase 5/7b: Render results (pause for MSRP comparison if mismatches exist)
        if (bulkState.msrpMismatches.length > 0) {
            // Default results toggle to MSRP when mismatches exist and any default chose 'current' (DB price)
            const hasCurrentDefaults = Array.from(bulkState.msrpChoices.values()).some(v => v === 'current');
            bulkState.resultsPricingMode = hasCurrentDefaults ? 'msrp' : 'reseller';
            bulkShowMsrpComparisonPanel();
        } else {
            bulkState.resultsPricingMode = 'reseller';
            bulkDisplayResults();
            scrollToPanel('bulkResultsPanel');
        }

    } catch (err) {
        console.error('[BulkSearch] RPC error:', err);
        // Show error using unmatched display slot
        bulkClearUnmatchedDisplay();
        const container = document.getElementById('bulkParsedRow');
        if (container) {
            const div = document.createElement('div');
            div.className = 'bulk-unmatched';
            div.style.cssText = 'margin-top:6px;padding:6px 8px;background:rgba(239,68,68,0.12);border:1px solid rgba(239,68,68,0.5);border-radius:4px;font-size:var(--font-size-xs);color:var(--color-error)';
            div.innerHTML = `<strong>Error loading products:</strong> ${err.message || 'Unknown error contacting Supabase'}`;
            container.appendChild(div);
        }
    } finally {
        // Hide progress bar
        if (progressEl) progressEl.classList.remove('visible');

        // Restore button
        if (loadBtn) {
            loadBtn.disabled = false;
            loadBtn.innerHTML = originalBtnHTML;
        }

        bulkState.isLoading = false;
    }
}

function bulkShowUnmatchedDisplay(list) {
    bulkClearUnmatchedDisplay();
    const container = document.getElementById('bulkParsedRow');
    if (!container) return;

    const div = document.createElement('div');
    div.className = 'bulk-unmatched';
    div.style.cssText = [
        'margin-top: 6px',
        'padding: 6px 8px',
        'background: rgba(239,68,68,0.08)',
        'border: 1px solid rgba(239,68,68,0.3)',
        'border-radius: 4px',
        'font-size: var(--font-size-xs)',
        'color: var(--color-error)',
        'line-height: 1.5',
    ].join('; ');

    div.innerHTML = `<strong>${list.length} MPN${list.length > 1 ? 's' : ''} not found:</strong> ${list.join(', ')}`;
    container.appendChild(div);
}

function bulkClearUnmatchedDisplay() {
    const els = document.querySelectorAll('.bulk-unmatched');
    els.forEach(el => el.remove());
}

function bulkClearSearch() {
    // Reset bulkState to initial values
    bulkState.products = [];
    bulkState.unmatchedMpns = [];
    bulkState.isLoading = false;
    bulkState.parsedSkus = [];
    bulkState.fileRows = [];
    bulkState.workbook = null;
    bulkState.parsedFileData = null;
    bulkState.fileName = null;
    bulkState.selectionMode = null;
    bulkState.hiddenColumns = new Set();
    bulkState.hiddenRows = new Set();
    bulkState.previewZoom = 55;
    bulkState.userManuallyZoomed = false;
    bulkState.userHasResized = false;
    bulkState.activeHiddenRowsDropdown = null;
    bulkState.selectedProductIndices.clear();
    bulkState.collapsedGroups.clear();
    bulkState.rawRpcRows.clear();
    bulkState.msrpMismatches = [];
    bulkState.msrpChoices = new Map();
    bulkState.resultsPricingMode = 'reseller';
    bulkState.resultsPage = 1;
    bulkState.detectedDistributor = null;
    bulkState.detectionSucceeded = false;
    state.currentDistributor = null;
    updateBulkDistributorBadges();
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    var tabsContainer = document.querySelector('.distributor-tabs');
    if (tabsContainer) tabsContainer.classList.remove('pulse-attention');

    // Reset collapsible sections to expanded
    if (!bulkState.collapsedSections) bulkState.collapsedSections = new Set();
    bulkState.collapsedSections.forEach(function(sectionId) {
        var header = document.querySelector('.bulk-collapsible-header[data-section="' + sectionId + '"]');
        var content = document.getElementById(sectionId);
        if (header) header.classList.remove('bulk-section-collapsed');
        if (content) content.classList.remove('bulk-section-hidden');
    });
    bulkState.collapsedSections.clear();

    // Hide MSRP comparison panel and re-edit button
    bulkHideMsrpComparisonPanel();
    const reEditBtn = document.getElementById('bulkMsrpReEditBtn');
    if (reEditBtn) reEditBtn.style.display = 'none';

    // Hide results panel
    const resultsPanel = document.getElementById('bulkResultsPanel');
    if (resultsPanel) resultsPanel.style.display = 'none';

    // Clear paste area
    const pasteArea = document.getElementById('bulkPasteArea');
    if (pasteArea) pasteArea.value = '';

    // Reset file input
    const fileInput = document.getElementById('bulkFileInput');
    if (fileInput) fileInput.value = '';

    // Hide spreadsheet preview
    bulkHideSpreadsheetPreview();

    // Hide parsed row
    const parsedRow = document.getElementById('bulkParsedRow');
    if (parsedRow) parsedRow.style.display = 'none';

    // Hide action bar
    const actionBar = document.getElementById('bulkActionBar');
    if (actionBar) actionBar.style.display = 'none';

    // Hide progress bar
    const progressEl = document.getElementById('bulkProgress');
    if (progressEl) progressEl.classList.remove('visible');

    // Clear unmatched display
    bulkClearUnmatchedDisplay();

    // Reset drop zone
    const dropZone = document.getElementById('bulkDropZone');
    if (dropZone) dropZone.classList.remove('has-file');
    const dropClearBtn = document.getElementById('bulkDropZoneClearBtn');
    if (dropClearBtn) dropClearBtn.classList.remove('visible');
    const dropStatus = document.getElementById('bulkDropZoneStatus');
    if (dropStatus) {
        dropStatus.textContent = '';
        dropStatus.style.color = '';
    }

    // Reset sheet selector
    const sheetSelect = document.getElementById('bulkSheetSelect');
    if (sheetSelect) {
        sheetSelect.innerHTML = '';
        sheetSelect.classList.remove('visible');
    }
    const sheetLabel = document.getElementById('bulkSheetSelectLabel');
    if (sheetLabel) sheetLabel.classList.remove('visible');

    // Disable mappings panel
    const mappingsPanel = document.getElementById('bulkMappingsPanel');
    if (mappingsPanel) mappingsPanel.classList.add('disabled');

    // Reset column dropdowns and row inputs
    bulkResetColumnDropdowns();
    const headerRowInput = document.getElementById('bulkHeaderRowInput');
    if (headerRowInput) headerRowInput.value = '1';
    const lastRowInput = document.getElementById('bulkLastRowInput');
    if (lastRowInput) lastRowInput.value = '';

    // Update parsed preview (will show empty state)
    bulkUpdateParsedPreview();

    bulkUpdateLoadButtonState();
    updateFooterStats();

    console.log('[BulkSearch] Cleared all bulk search state');
}

// =====================================================
// BULK SEARCH — Phase 5: Results Rendering
// =====================================================

function bulkDisplayResults() {
    const tbody = document.getElementById('bulkProductsBody');
    const countEl = document.getElementById('bulkResultsCount');
    const badgesEl = document.getElementById('bulkMfrBadges');
    const emptyState = document.getElementById('bulkEmptyState');
    const resultsPanel = document.getElementById('bulkResultsPanel');

    // Sync results pricing toggle active state
    const resultsToggle = document.getElementById('bulkResultsPricingToggle');
    if (resultsToggle) {
        resultsToggle.querySelectorAll('.pricing-toggle-btn').forEach(btn => {
            btn.classList.toggle('active', btn.getAttribute('data-price') === bulkState.resultsPricingMode);
        });
    }

    // Show/hide re-edit MSRP button
    const reEditBtn = document.getElementById('bulkMsrpReEditBtn');
    if (reEditBtn) {
        reEditBtn.style.display = bulkState.msrpMismatches.length > 0 ? '' : 'none';
    }

    if (bulkState.products.length === 0) {
        tbody.innerHTML = '';
        countEl.textContent = '0 products';
        badgesEl.innerHTML = '';
        emptyState.style.display = 'flex';
        resultsPanel.style.display = 'none';
        if (reEditBtn) reEditBtn.style.display = 'none';
        const pagination = document.getElementById('bulkPagination');
        if (pagination) pagination.style.display = 'none';
        return;
    }

    emptyState.style.display = 'none';
    resultsPanel.style.display = '';
    const foundCount = bulkState.products.filter(p => !p.not_found).length;
    countEl.textContent = `${foundCount} product${foundCount !== 1 ? 's' : ''}`;

    // Group by manufacturer (exclude not-found products)
    const grouped = {};
    bulkState.products.forEach((p) => {
        if (p.not_found) return;
        const mfr = p.manufacturer || 'Unknown';
        if (!grouped[mfr]) grouped[mfr] = [];
        grouped[mfr].push(p);
    });

    // Render manufacturer badges (ordered by first appearance in spreadsheet)
    const mfrList = Object.keys(grouped);
    badgesEl.innerHTML = mfrList
        .map(mfr => `<span class="bulk-mfr-badge">${mfr}</span>`)
        .join('');

    // Pagination: compute page range
    const startIdx = (bulkState.resultsPage - 1) * bulkState.resultsPerPage;
    const endIdx = startIdx + bulkState.resultsPerPage;

    // Build a flat list of global indices for found products (preserving group order)
    let globalCounter = 0;

    // Render table with collapsible groups
    let html = '';
    mfrList.forEach(mfr => {
        const products = grouped[mfr];
        const isCollapsed = bulkState.collapsedGroups.has(mfr);

        // Determine which products in this group fall on the current page
        const pageProducts = [];
        products.forEach(p => {
            if (p.not_found) return;
            if (globalCounter >= startIdx && globalCounter < endIdx) {
                pageProducts.push(p);
            }
            globalCounter++;
        });

        // Only render this manufacturer group if it has products on this page
        if (pageProducts.length === 0) return;

        // Manufacturer group divider — show TOTAL count for the group, not just page slice
        html += `<tr><td class="bulk-mfr-group-divider ${isCollapsed ? 'bulk-collapsed' : ''}" data-mfr="${mfr}" onclick="bulkToggleGroup('${mfr.replace(/'/g, "\\'")}')">` +
            `<svg class="bulk-collapse-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>` +
            `${mfr} <span class="bulk-mfr-group-count">${products.length} product${products.length !== 1 ? 's' : ''}</span>` +
            `</td></tr>`;

        // Product rows — only the page slice
        pageProducts.forEach(p => {
            const globalIdx = bulkState.products.indexOf(p);
            const isSelected = bulkState.selectedProductIndices.has(globalIdx);
            const hiddenClass = isCollapsed ? 'bulk-hidden' : '';
            const escapedMfr = mfr.replace(/"/g, '&quot;');

            const price = bulkState.resultsPricingMode === 'reseller' && p.resellerPrice != null ? p.resellerPrice : p.msrp;
            const msrpIndicator = p._msrpAdjusted
                ? (bulkState.resultsPricingMode === 'msrp'
                    ? `<span class="bulk-msrp-indicator bulk-msrp-${p._msrpDirection}">${p._msrpDirection === 'down' ? '▼' : '▲'}</span>`
                    : `<span class="bulk-msrp-dot bulk-msrp-${p._msrpDirection}">●</span>`)
                : '';
            html += `<tr class="bulk-product-row ${isSelected ? 'bulk-selected' : ''} ${hiddenClass}" data-index="${globalIdx}" data-mfr="${escapedMfr}">` +
                `<td class="bulk-col-checkbox"><input type="checkbox" ${isSelected ? 'checked' : ''} onchange="bulkToggleProductSelection(${globalIdx})"></td>` +
                `<td class="bulk-col-part">${(p.mpn || '')}</td>` +
                `<td class="bulk-col-desc" title="${(p.description || '').replace(/"/g, '&quot;')}">${p.description || ''}</td>` +
                `<td class="bulk-col-price">${bulkFormatPrice(price)}</td>` +
                `<td class="bulk-col-indicator">${msrpIndicator}</td>` +
                `<td class="bulk-col-action"><button class="bulk-info-btn" onclick="bulkShowProductInfo(${globalIdx})">i</button></td></tr>`;
        });
    });

    tbody.innerHTML = html;
    bulkUpdateResultsSelectionUI();
    bulkUpdatePagination();

    // Auto-expand results panel to fit content (up to 50 rows worth of height)
    // Remove any previously set max-height so panel expands naturally
    resultsPanel.style.maxHeight = 'none';
    resultsPanel.style.overflow = 'visible';

    if (!bulkState.msrpMismatches || bulkState.msrpMismatches.length === 0) {
        document.getElementById('bulkResultsPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function bulkToggleGroup(mfr) {
    if (bulkState.collapsedGroups.has(mfr)) {
        bulkState.collapsedGroups.delete(mfr);
    } else {
        bulkState.collapsedGroups.add(mfr);
    }

    const divider = document.querySelector(`.bulk-mfr-group-divider[data-mfr="${mfr}"]`);
    const rows = document.querySelectorAll(`.bulk-product-row[data-mfr="${mfr}"]`);

    if (bulkState.collapsedGroups.has(mfr)) {
        if (divider) divider.classList.add('bulk-collapsed');
        rows.forEach(row => row.classList.add('bulk-hidden'));
    } else {
        if (divider) divider.classList.remove('bulk-collapsed');
        rows.forEach(row => row.classList.remove('bulk-hidden'));
    }
}

function bulkToggleProductSelection(index) {
    if (bulkState.selectedProductIndices.has(index)) {
        bulkState.selectedProductIndices.delete(index);
    } else {
        bulkState.selectedProductIndices.add(index);
    }

    const row = document.querySelector(`.bulk-product-row[data-index="${index}"]`);
    if (row) {
        row.classList.toggle('bulk-selected', bulkState.selectedProductIndices.has(index));
    }

    bulkUpdateResultsSelectionUI();
    updateFooterStats();
}

function bulkToggleSelectAll() {
    const checkbox = document.getElementById('bulkSelectAll');

    if (checkbox.checked) {
        bulkState.products.forEach((p, i) => {
            if (!p.not_found) bulkState.selectedProductIndices.add(i);
        });
    } else {
        bulkState.selectedProductIndices.clear();
    }

    bulkDisplayResults();
    updateFooterStats();
}

function bulkUpdateResultsSelectionUI() {
    const addBtn = document.getElementById('bulkAddToQueueBtn');
    const selectAllCheckbox = document.getElementById('bulkSelectAll');
    const selectableCount = bulkState.products.filter(p => !p.not_found).length;

    if (addBtn) addBtn.disabled = bulkState.selectedProductIndices.size === 0;
    if (selectAllCheckbox) {
        selectAllCheckbox.checked = bulkState.selectedProductIndices.size === selectableCount && selectableCount > 0;
        selectAllCheckbox.indeterminate = bulkState.selectedProductIndices.size > 0 && bulkState.selectedProductIndices.size < selectableCount;
    }
}

function bulkAddSelectedToQueue() {
    const selected = [...bulkState.selectedProductIndices].sort((a, b) => a - b).map(i => bulkState.products[i]).filter(Boolean);

    let added = 0;
    let skipped = 0;
    selected.forEach(p => {
        if (!bulkState.queuedProducts.find(q => q.mpn === p.mpn || q.vendorPartNumber === p.mpn)) {
            const mpnKey = (p.mpn || '').toUpperCase();
            const rawRow = bulkState.rawRpcRows.get(mpnKey);
            bulkState.queuedProducts.push({ ...p, vendorPartNumber: p.mpn, _rawRpcRow: rawRow || null, customerDiscount: 0 });
            added++;
        } else {
            skipped++;
        }
    });

    bulkState.selectedProductIndices.clear();
    bulkDisplayResults();
    updateQueueUI();
    updateFooterStats();
    document.querySelector('.queue-panel').scrollIntoView({ behavior: 'smooth', block: 'start' });

    if (skipped > 0 && added === 0) {
        bulkShowToast('bulkQueueToast', 'All items already in queue', 'bulkResultsPanel');
    } else if (skipped > 0) {
        bulkShowToast('bulkQueueToast', `Added ${added} item${added !== 1 ? 's' : ''}, ${skipped} duplicate${skipped !== 1 ? 's' : ''} skipped`, 'bulkResultsPanel');
    } else if (added > 0) {
        bulkShowToast('bulkQueueToast', `Added ${added} item${added !== 1 ? 's' : ''} to queue`, 'bulkResultsPanel');
    }

    console.log(`[BulkSearch] Added ${added} to queue, ${skipped} duplicates skipped`);
}

function bulkShowProductInfo(index) {
    const product = bulkState.products[index];
    if (!product || product.not_found) return;

    // Look up the raw RPC row using the product's MPN
    const mpnKey = (product.mpn || '').toUpperCase();
    const rawRow = bulkState.rawRpcRows.get(mpnKey);

    let mapped;
    switch (state.currentDistributor) {
        case 'tdsynnex':
            if (rawRow) {
                mapped = mapTDSynnexProduct(rawRow);
            } else {
                mapped = { ...product, vendorPartNumber: product.mpn, _source: 'tdsynnex' };
            }
            break;
        case 'adi':
            if (rawRow) {
                mapped = mapADIGlobalProduct(rawRow);
            } else {
                mapped = { ...product, vendorPartNumber: product.mpn, adiSku: product.vpn, _source: 'adi' };
            }
            break;
        case 'ingram':
        default:
            if (rawRow) {
                // Map Ingram RPC row to the format showProductDetails expects
                // Product info from DB, pricing from API (fetched by showProductDetails)
                mapped = {
                    ingramPartNumber: rawRow.ingram_part_number || product._fileVpn || product.mpn || '',
                    vendorPartNumber: rawRow.vendor_part_number || '',
                    vendorName: rawRow.manufacturer || rawRow.vendor_name || '',
                    description: rawRow.description_line_1 || rawRow.description || '',
                    extraDescription: [rawRow.description_line_1, rawRow.description_line_2].filter(Boolean).join(' '),
                    category: rawRow.level_1_name || rawRow.category || '',
                    subCategory: rawRow.level_2_name || '',
                    productType: rawRow.media_type || '',
                    replacementSku: rawRow.substitute_part_number || '',
                    upcCode: rawRow.upc_code || '',
                    retailPrice: rawRow.retail_price,
                    customerPrice: rawRow.customer_price,
                    pricingData: null,
                    _source: 'ingram'
                };
            } else {
                mapped = { ...product, vendorPartNumber: product.mpn, ingramPartNumber: product._fileVpn || product.vpn || product.mpn, _source: 'ingram' };
            }
            break;
    }

    // Temporarily place the fully mapped product into state.currentProducts so showProductDetails can find it
    const tempIndex = state.currentProducts.length;
    state.currentProducts.push(mapped);
    showProductDetails(tempIndex);
}

function bulkFormatPrice(value) {
    if (value === null || value === undefined || value === '') return '\u2014';
    const num = parseFloat(value);
    if (isNaN(num)) return '\u2014';
    return '$' + num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// =====================================================
// BULK SEARCH — Results Pricing Toggle
// =====================================================

function bulkSetResultsPricingMode(mode) {
    bulkState.resultsPricingMode = mode;

    // Update only the results toggle button states (not the queue toggle)
    const toggle = document.getElementById('bulkResultsPricingToggle');
    if (toggle) {
        toggle.querySelectorAll('.pricing-toggle-btn').forEach(btn => {
            btn.classList.toggle('active', btn.getAttribute('data-price') === mode);
        });
    }

    // Re-render results with new pricing mode
    bulkDisplayResults();
}

// =====================================================
// Phase 7b Step 3: MSRP Comparison Panel
// =====================================================

/**
 * Populates and shows the MSRP comparison panel.
 * Called from bulkLoadProducts() when mismatches are detected.
 */
function bulkShowMsrpComparisonPanel() {
    const panel = document.getElementById('bulkMsrpComparisonPanel');
    const tbody = document.getElementById('bulkMsrpComparisonBody');
    if (!panel || !tbody) return;

    tbody.innerHTML = '';

    bulkState.msrpMismatches.forEach(item => {
        const mpnKey = item.mpn.toUpperCase();
        const choice = bulkState.msrpChoices.get(mpnKey) || 'quote';
        const diff = item.difference; // dbMsrp - fileMsrp
        const isDecrease = diff < 0; // current DB price is lower
        const absDiff = Math.abs(diff);

        const changeArrow = isDecrease ? '\u25BC' : '\u25B2';
        const changeClass = isDecrease ? 'bulk-msrp-change-down' : 'bulk-msrp-change-up';
        const rowClass = choice === 'current' ? 'bulk-msrp-row-selected-current' : 'bulk-msrp-row-selected-quote';

        // Truncate description
        const desc = (item.description || '').length > 50
            ? item.description.substring(0, 47) + '...'
            : (item.description || '');

        const tr = document.createElement('tr');
        tr.id = 'bulkMsrpRow_' + mpnKey;
        tr.className = rowClass;
        tr.innerHTML = `
            <td title="${item.mpn}">${item.mpn}</td>
            <td title="${item.description || ''}">${desc}</td>
            <td>${bulkFormatPrice(item.fileMsrp)}</td>
            <td>${bulkFormatPrice(item.dbMsrp)}</td>
            <td class="${changeClass}">${bulkFormatPrice(absDiff).replace('$', '$\u200B')} ${changeArrow}</td>
            <td>
                <div class="bulk-msrp-radio-group">
                    <label><input type="radio" name="bulkMsrpChoice_${mpnKey}" value="quote"
                        ${choice === 'quote' ? 'checked' : ''}
                        onchange="bulkMsrpToggleChoice('${mpnKey}', 'quote')">Q</label>
                    <label><input type="radio" name="bulkMsrpChoice_${mpnKey}" value="current"
                        ${choice === 'current' ? 'checked' : ''}
                        onchange="bulkMsrpToggleChoice('${mpnKey}', 'current')">C</label>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });

    // Hide results panel while comparison is active
    const resultsPanel = document.getElementById('bulkResultsPanel');
    if (resultsPanel) resultsPanel.style.display = 'none';

    panel.style.display = 'flex';
    document.getElementById('bulkMsrpComparisonPanel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/**
 * Sets all MSRP choices to 'quote' or 'current'.
 */
function bulkMsrpChooseAll(choice) {
    bulkState.msrpMismatches.forEach(item => {
        const mpnKey = item.mpn.toUpperCase();
        bulkState.msrpChoices.set(mpnKey, choice);

        // Update radio button
        const radio = document.querySelector(`input[name="bulkMsrpChoice_${mpnKey}"][value="${choice}"]`);
        if (radio) radio.checked = true;

        // Update row styling
        const row = document.getElementById('bulkMsrpRow_' + mpnKey);
        if (row) {
            row.className = choice === 'current' ? 'bulk-msrp-row-selected-current' : 'bulk-msrp-row-selected-quote';
        }
    });
}

/**
 * Handles individual radio button change.
 */
function bulkMsrpToggleChoice(mpnKey, choice) {
    bulkState.msrpChoices.set(mpnKey, choice);

    const row = document.getElementById('bulkMsrpRow_' + mpnKey);
    if (row) {
        row.className = choice === 'current' ? 'bulk-msrp-row-selected-current' : 'bulk-msrp-row-selected-quote';
    }
}

/**
 * Applies MSRP choices to bulkState.products and proceeds to results.
 */
function bulkApplyMsrpChoices() {
    bulkState.products.forEach(p => {
        if (p._dbMsrp != null && p._fileMsrp != null && p._dbMsrp !== p._fileMsrp) {
            const mpnKey = p.mpn.toUpperCase();
            const choice = bulkState.msrpChoices.get(mpnKey) || 'quote';

            if (choice === 'current') {
                p.msrp = p._dbMsrp;
                p._msrpAdjusted = true;
                p._msrpDirection = (p._dbMsrp < p._fileMsrp) ? 'down' : 'up';
            } else {
                p.msrp = p._fileMsrp;
                p._msrpAdjusted = false;
            }
        }
    });

    // Set results pricing toggle default based on choices
    const hasCurrentChoice = Array.from(bulkState.msrpChoices.values()).some(v => v === 'current');
    bulkState.resultsPricingMode = hasCurrentChoice ? 'msrp' : 'reseller';

    bulkHideMsrpComparisonPanel();
    bulkDisplayResults();
    scrollToPanel('bulkResultsPanel');

    // Show re-edit button if mismatches exist
    const reEditBtn = document.getElementById('bulkMsrpReEditBtn');
    if (reEditBtn) {
        reEditBtn.style.display = bulkState.msrpMismatches.length > 0 ? '' : 'none';
    }
}

/**
 * Hides the MSRP comparison panel.
 */
function bulkHideMsrpComparisonPanel() {
    const panel = document.getElementById('bulkMsrpComparisonPanel');
    if (panel) panel.style.display = 'none';
}

// Bulk spreadsheet preview vertical resize
function initBulkPreviewResize() {
    const handle = document.getElementById('bulkPreviewResizeHandle');
    const target = document.getElementById('bulkSpreadsheetPreview');

    if (!handle || !target) return;

    let isResizingPreview = false;
    let previewResizeStartY = 0;
    let previewResizeStartHeight = 0;

    handle.addEventListener('mousedown', (e) => {
        isResizingPreview = true;
        previewResizeStartY = e.clientY;
        previewResizeStartHeight = target.offsetHeight;
        document.body.style.cursor = 'ns-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizingPreview) return;

        const deltaY = e.clientY - previewResizeStartY;
        const newHeight = Math.max(80, Math.min(800, previewResizeStartHeight + deltaY));
        target.style.height = newHeight + 'px';
        bulkState.userHasResized = true;
    });

    document.addEventListener('mouseup', () => {
        if (isResizingPreview) {
            isResizingPreview = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// Bulk results panel vertical resize
function initBulkResultsResize() {
    const handle = document.getElementById('bulkResultsResizeHandle');
    const target = document.getElementById('bulkResultsPanel');

    if (!handle || !target) return;

    let isResizingResults = false;
    let resultsResizeStartY = 0;
    let resultsResizeStartHeight = 0;

    handle.addEventListener('mousedown', (e) => {
        isResizingResults = true;
        resultsResizeStartY = e.clientY;
        resultsResizeStartHeight = target.offsetHeight;
        document.body.style.cursor = 'ns-resize';
        document.body.style.userSelect = 'none';
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizingResults) return;

        const deltaY = e.clientY - resultsResizeStartY;
        const newHeight = Math.max(100, Math.min(2000, resultsResizeStartHeight + deltaY));
        target.style.maxHeight = newHeight + 'px';
        target.style.overflow = 'hidden';
    });

    document.addEventListener('mouseup', () => {
        if (isResizingResults) {
            isResizingResults = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// =====================================================
// BULK SEARCH — Phase 8: Pagination
// =====================================================

function bulkChangePage(delta) {
    const totalProducts = bulkState.products.filter(p => !p.not_found).length;
    const totalPages = Math.ceil(totalProducts / bulkState.resultsPerPage);
    const newPage = bulkState.resultsPage + delta;
    if (newPage < 1 || newPage > totalPages) return;
    bulkState.resultsPage = newPage;
    bulkDisplayResults();
}

function bulkUpdatePagination() {
    const pagination = document.getElementById('bulkPagination');
    const totalProducts = bulkState.products.filter(p => !p.not_found).length;
    const totalPages = Math.ceil(totalProducts / bulkState.resultsPerPage);

    if (totalPages <= 1) {
        if (pagination) pagination.style.display = 'none';
        return;
    }

    if (pagination) {
        pagination.style.display = 'flex';
        document.getElementById('bulkPageInfo').textContent = 'Page ' + bulkState.resultsPage + ' of ' + totalPages;
        document.getElementById('bulkPrevPage').disabled = bulkState.resultsPage <= 1;
        document.getElementById('bulkNextPage').disabled = bulkState.resultsPage >= totalPages;
    }
}

// =====================================================
// BULK SEARCH — Phase 8.6: Collapsible Section Headers
// =====================================================

// Track collapsed state
bulkState.collapsedSections = new Set();

function bulkToggleSection(sectionId) {
    var header = document.querySelector('.bulk-collapsible-header[data-section="' + sectionId + '"]');
    var content = document.getElementById(sectionId);
    if (!header || !content) return;

    if (bulkState.collapsedSections.has(sectionId)) {
        bulkState.collapsedSections.delete(sectionId);
        header.classList.remove('bulk-section-collapsed');
        content.classList.remove('bulk-section-hidden');

        // Re-expand textarea after section becomes visible
        if (sectionId === 'bulkParsedContent') {
            requestAnimationFrame(function() {
                var el = document.getElementById('bulkParsedSkusEditable');
                if (!el || !el.value) return;
                el.style.height = 'auto';
                el.style.height = Math.max(48, el.scrollHeight + 4) + 'px';
            });
        }
    } else {
        bulkState.collapsedSections.add(sectionId);
        header.classList.add('bulk-section-collapsed');
        content.classList.add('bulk-section-hidden');
    }
}

// =====================================================
// BULK SEARCH — Phase 8.7: Cmd/Ctrl + Scroll Wheel Zoom on Spreadsheet Preview
// =====================================================
// When user holds Cmd (Mac) or Ctrl (Windows/Linux) and scrolls the mouse wheel
// over the spreadsheet preview area, zoom in/out. Without modifier key, normal
// scrolling behavior is preserved. Reuses existing bulkZoomPreview() which handles
// min/max clamping (25%-150%) and sets userManuallyZoomed = true.

function initBulkScrollWheelZoom() {
    const previewContainer = document.getElementById('bulkPreviewContainer');
    if (!previewContainer) return;

    previewContainer.addEventListener('wheel', function(e) {
        // Only intercept when Cmd (Mac) or Ctrl (Windows/Linux) is held
        if (e.ctrlKey || e.metaKey) {
            e.preventDefault();
            // Scroll up (negative deltaY) = zoom in (+1), scroll down = zoom out (-1)
            const direction = e.deltaY < 0 ? 1 : -1;
            bulkZoomPreview(direction);
        }
    }, { passive: false });
}

// ========== BULK DISTRIBUTOR AUTO-DETECTION (Phase 9.2 Step 2) ==========

/**
 * Scan the first ~20 rows of bulkState.fileRows for distributor keywords
 * and auto-select the matching distributor tab if found.
 * Returns the detected distributor string ('ingram', 'tdsynnex') or null.
 */
function bulkDetectDistributor() {
    console.log('[BulkDetect] bulkDetectDistributor() called');

    if (!bulkState.fileRows || bulkState.fileRows.length === 0) {
        console.log('[BulkDetect] No fileRows available, returning null');
        return null;
    }

    console.log('[BulkDetect] Starting distributor detection, fileRows:', bulkState.fileRows.length, 'rows');

    // Scan ALL rows, return immediately on first match
    for (let r = 0; r < bulkState.fileRows.length; r++) {
        const row = bulkState.fileRows[r];
        if (!row) continue;

        // Concatenate all cells in this row into a single lowercase string
        let rowText = '';
        for (let c = 0; c < row.length; c++) {
            const cellVal = String(row[c] || '').trim().toLowerCase();
            if (cellVal) {
                rowText += ' ' + cellVal;
            }
        }

        if (!rowText) continue;

        console.log('[BulkDetect] Row', r, ':', rowText.substring(0, 100));

        // Check for distributor keywords — email domains first (most specific), then word-boundary terms
        let detected = null;

        if (/tdsynnex\.com/.test(rowText) || /td\s*synnex/.test(rowText) || /\bsynnex\b/.test(rowText)) {
            detected = 'tdsynnex';
        } else if (/ingrammicro\.com/.test(rowText) || /\bingram\b/.test(rowText)) {
            detected = 'ingram';
        } else if (/adiglobal\.com/.test(rowText) || /\badi\s*global\b/.test(rowText) || /\badi\b/.test(rowText)) {
            detected = 'adi';
        }

        if (detected) {
            const displayName = DISTRIBUTORS[detected] ? DISTRIBUTORS[detected].name : detected;
            console.log('[BulkDetect] MATCH FOUND:', detected, 'in row', r, '- selecting distributor:', displayName);
            bulkState.detectedDistributor = detected;
            bulkState.detectionSucceeded = true;
            selectDistributor(detected);
            bulkUpdateLoadButtonState();
            showStatus('Auto-detected distributor: ' + displayName, 'success');
            return detected;
        }
    }

    console.log('[BulkDetect] No distributor detected after scanning all', bulkState.fileRows.length, 'rows');
    bulkState.detectedDistributor = null;
    bulkState.detectionSucceeded = false;
    bulkUpdateLoadButtonState();
    return null;
}

// ========== BULK AUTO-DETECT HEADER ROW (Phase 9.2 Step 3) ==========

function bulkAutoDetectHeaderRow() {
    // Returns a Promise that resolves to the 0-based header row index

    var detected = state.currentDistributor; // already set by bulkDetectDistributor()

    console.log('[BulkAutoMap] Starting header row detection for distributor:', detected);

    // Step 1: Query Supabase for stored header_row
    return fetch(SUPABASE_URL + '/rest/v1/bulk_column_mapping_rules?distributor=eq.' + encodeURIComponent(detected) + '&select=header_row', {
        headers: {
            'apikey': SUPABASE_ANON_KEY,
            'Authorization': 'Bearer ' + SUPABASE_ANON_KEY
        }
    })
    .then(function(res) { return res.json(); })
    .then(function(data) {
        var storedRow = (data && data[0] && data[0].header_row != null) ? data[0].header_row : null;
        console.log('[BulkAutoMap] Supabase stored header_row for', detected, ':', storedRow);

        if (storedRow !== null && storedRow >= 1) {
            // Validate: check if that row has header-like content
            var rowIndex = storedRow - 1; // convert to 0-based
            if (rowIndex < bulkState.fileRows.length && bulkIsHeaderRow(rowIndex)) {
                console.log('[BulkAutoMap] Stored header row', storedRow, 'validated successfully');
                return rowIndex;
            }
            console.log('[BulkAutoMap] Stored header row', storedRow, 'failed validation, falling back to heuristic');
        }

        // Step 2: Heuristic scan
        return bulkHeuristicHeaderRowScan();
    })
    .catch(function(err) {
        console.warn('[BulkAutoMap] Supabase query failed, using heuristic:', err);
        return bulkHeuristicHeaderRowScan();
    });
}

function bulkIsHeaderRow(rowIndex) {
    var row = bulkState.fileRows[rowIndex];
    if (!row) return false;
    var score = bulkScoreHeaderRow(row);
    return score >= 2; // at least 2 keyword matches
}

function bulkScoreHeaderRow(row) {
    // All known header keywords (lowercase)
    var keywords = [
        // MPN
        'item number', 'mpn', 'part number', 'mfg part', 'manufacturer part', 'sku', 'part #', 'part no', 'mfr. part',
        // QTY
        'qty', 'quantity',
        // Price
        'price', 'reseller', 'cost', 'dealer price', 'our price', 'unit price', 'unit cost', 'customer price', 'contract price',
        // VPN
        'vpn', 'vendor part', 'ingram part', 'vendor name',
        // MSRP
        'msrp', 'list price', 'retail price', 'suggested retail',
        // Other common headers
        'description', 'availability', 'rebate', 'quote line', 'spa ref'
    ];

    var score = 0;
    var matched = [];
    for (var c = 0; c < row.length; c++) {
        var cellVal = String(row[c] || '').trim().toLowerCase();
        if (!cellVal) continue;
        for (var k = 0; k < keywords.length; k++) {
            if (cellVal.indexOf(keywords[k]) !== -1) {
                score++;
                matched.push(cellVal);
                break; // one match per cell is enough
            }
        }
    }
    console.log('[BulkAutoMap] Row scored', score, 'matches:', matched.join(', '));
    return score;
}

function bulkAutoDetectLastRow(headerRowIndex) {
    // Detect the last data row using column-consistency analysis.
    // Real data rows populate roughly the same columns. Summary/totals rows
    // leave significant gaps where data rows have values.
    var rows = bulkState.fileRows;
    if (!rows || rows.length === 0) return headerRowIndex;

    var mpnColSelect = document.getElementById('bulkColumnSelect');
    var mpnColIdx = (mpnColSelect && mpnColSelect.value !== '') ? parseInt(mpnColSelect.value) : -1;

    // Step 1: Build reference pattern from first 3 data rows after header.
    // Collect which column indices are populated across those rows (union).
    var referencePopulated = {};  // column index -> true
    var refRowCount = 0;
    var scanEnd = Math.min(headerRowIndex + 4, rows.length - 1);
    for (var r = headerRowIndex + 1; r <= scanEnd; r++) {
        var row = rows[r];
        if (!row) continue;
        var popCount = 0;
        for (var c = 0; c < row.length; c++) {
            if (String(row[c] || '').trim() !== '') popCount++;
        }
        if (popCount >= 5) {
            for (var c2 = 0; c2 < row.length; c2++) {
                if (String(row[c2] || '').trim() !== '') {
                    referencePopulated[c2] = true;
                }
            }
            refRowCount++;
        }
        if (refRowCount >= 3) break;
    }

    // Count reference columns
    var refCols = Object.keys(referencePopulated);
    var refSize = refCols.length;

    // Fallback if no reference rows found (very sparse file)
    if (refRowCount === 0) {
        for (var r2 = rows.length - 1; r2 > headerRowIndex; r2--) {
            var row2 = rows[r2];
            if (!row2) continue;
            var cnt = 0;
            for (var c3 = 0; c3 < row2.length; c3++) {
                if (String(row2[c3] || '').trim() !== '') cnt++;
            }
            if (cnt >= 3) return r2;
        }
        return headerRowIndex;
    }

    // Step 2: Scan backwards from bottom, test each candidate row
    var minThreshold = Math.max(3, Math.floor(refSize * 0.4));

    for (var r3 = rows.length - 1; r3 > headerRowIndex; r3--) {
        var row3 = rows[r3];
        if (!row3) continue;

        // Count populated cells in this row
        var populated = 0;
        for (var c4 = 0; c4 < row3.length; c4++) {
            if (String(row3[c4] || '').trim() !== '') populated++;
        }

        // Quick skip: too few cells populated
        if (populated < minThreshold) continue;

        // Column consistency: how many reference columns does this row leave blank?
        var blanksInRefCols = 0;
        for (var c5 = 0; c5 < refCols.length; c5++) {
            var colIdx = parseInt(refCols[c5]);
            var cellVal = (row3.length > colIdx) ? String(row3[colIdx] || '').trim() : '';
            if (cellVal === '') blanksInRefCols++;
        }

        var blankRatio = blanksInRefCols / refSize;

        // If row is missing >40% of columns that data rows populate → summary row, skip
        if (blankRatio > 0.40) {
            console.log('[BulkAutoMap] Skipping row ' + (r3 + 1) + ' as summary (blank ratio: ' + Math.round(blankRatio * 100) + '%, ' + blanksInRefCols + '/' + refSize + ' ref cols blank)');
            continue;
        }

        // MPN column gate: if MPN mapped and this row has no MPN, skip
        if (mpnColIdx >= 0) {
            var mpnVal = (row3.length > mpnColIdx) ? String(row3[mpnColIdx] || '').trim() : '';
            if (mpnVal === '') {
                console.log('[BulkAutoMap] Skipping row ' + (r3 + 1) + ' — no MPN value');
                continue;
            }
        }

        console.log('[BulkAutoMap] Last data row detected: ' + (r3 + 1) + ' (blank ratio: ' + Math.round(blankRatio * 100) + '%)');
        return r3;
    }

    // No data rows found
    return headerRowIndex;
}

// ========== BULK COLUMN MAPPING RULES FETCH (Phase 9.2 Step 4) ==========

function bulkFetchMappingRules() {
    // Fetch both distributor-specific and universal rules from Supabase
    // Returns a Promise resolving to { mpn: [...], qty: [...], price: [...], vpn: [...], msrp: [...] }
    var distributor = state.currentDistributor;
    var dbDistributorName = distributor === 'adi' ? 'adiglobal' : distributor;

    console.log('[BulkAutoMap] Fetching mapping rules for distributor:', dbDistributorName);

    // Query both the distributor row and universal row in one request
    var url = SUPABASE_URL + '/rest/v1/bulk_column_mapping_rules?distributor=in.(' +
        encodeURIComponent(dbDistributorName) + ',universal)&select=distributor,mpn,qty,price,vpn,msrp';

    return fetch(url, {
        headers: {
            'apikey': SUPABASE_ANON_KEY,
            'Authorization': 'Bearer ' + SUPABASE_ANON_KEY
        }
    })
    .then(function(res) { return res.json(); })
    .then(function(rows) {
        // Separate distributor-specific and universal rows
        var distRow = null;
        var univRow = null;
        for (var i = 0; i < rows.length; i++) {
            if (rows[i].distributor === dbDistributorName) distRow = rows[i];
            if (rows[i].distributor === 'universal') univRow = rows[i];
        }

        // Per-field fallback: use distributor keywords if present, otherwise universal (never merge both)
        var fields = ['mpn', 'qty', 'price', 'vpn', 'msrp'];
        var merged = {};
        for (var f = 0; f < fields.length; f++) {
            var field = fields[f];
            var keywords = [];
            if (distRow && distRow[field]) {
                // Distributor has keywords for this field — use ONLY those
                keywords = distRow[field].split(',').map(function(k) { return k.trim().toLowerCase(); }).filter(Boolean);
            } else if (univRow && univRow[field]) {
                // Distributor field empty — fall back to universal keywords
                keywords = univRow[field].split(',').map(function(k) { return k.trim().toLowerCase(); }).filter(Boolean);
            }
            console.log('[BulkAutoMap] Using ' + (distRow && distRow[field] ? 'distributor' : 'universal') + ' keywords for ' + field + ':', keywords);
            merged[field] = keywords;
        }

        console.log('[BulkAutoMap] Merged mapping rules:', merged);
        return merged;
    })
    .catch(function(err) {
        console.warn('[BulkAutoMap] Failed to fetch mapping rules, using hardcoded defaults:', err);
        // Return hardcoded defaults as fallback
        return {
            mpn: ['item number', 'mpn', 'part number', 'mfg part', 'manufacturer part', 'mfr part', 'mfr.', 'sku', 'part #', 'part no'],
            qty: ['qty', 'quantity'],
            price: ['reseller price', 'reseller', 'dealer price', 'our price', 'unit price', 'unit cost', 'customer price', 'contract price', 'wholesale price', 'net price'],
            vpn: ['vpn', 'vendor part', 'ingram part'],
            msrp: ['msrp', 'list price', 'list', 'retail price', 'suggested retail']
        };
    });
}

function bulkHeuristicHeaderRowScan() {
    console.log('[BulkAutoMap] Running heuristic header row scan');
    var bestRow = 0; // default to first row (0-based)
    var bestScore = 0;
    var maxScan = Math.min(30, bulkState.fileRows.length);

    for (var r = 0; r < maxScan; r++) {
        var row = bulkState.fileRows[r];
        if (!row) continue;
        var score = bulkScoreHeaderRow(row);
        if (score > bestScore) {
            bestScore = score;
            bestRow = r;
        }
    }

    if (bestScore >= 2) {
        console.log('[BulkAutoMap] Heuristic found header at row', bestRow + 1, '(0-based:', bestRow, ') with score', bestScore);
        return bestRow;
    }

    console.log('[BulkAutoMap] Heuristic found no strong header match, defaulting to row 1');
    return 0; // default to first row
}