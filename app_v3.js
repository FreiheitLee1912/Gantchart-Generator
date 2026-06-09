/**
 * Gantt Chart Generator - CSV to PPT-Ready Gantt Chart
 * Parses Jira CSV exports and renders interactive Gantt charts
 */

// ========================================
// Global State
// ========================================
const state = {
    files: [],
    tasks: [],
    filteredTasks: [],
    viewMode: 'month', // month, quarter, year
    groupBy: 'none',   // default to no grouping
    selectedProject: 'all',
    theme: 'light',
    dayWidth: 4, // pixels per day
    customStart: null,  // user-set timeline start (Date or null)
    customEnd: null,    // user-set timeline end (Date or null)
    hideSummaryColumn: true,
};

// ========================================
// Task Type Configuration
// ========================================
const TYPE_CONFIG = {
    'Milestone':   { color: 'var(--milestone-color)',   hex: '#f4a6a6', label: '*' },
    'Goal':        { color: 'var(--task-color)',         hex: '#003b75', label: 'Goal' },
    'Review':      { color: 'var(--task-color)',         hex: '#003b75', label: 'Rev' },
    'Event':       { color: 'var(--task-color)',         hex: '#003b75', label: 'Evt' },
    'Task':        { color: 'var(--task-color)',         hex: '#003b75', label: 'Task' },
    'Sub task':    { color: 'var(--task-color)',         hex: '#003b75', label: 'Sub' },
    'Validation':  { color: 'var(--task-color)',         hex: '#003b75', label: 'Val' },
    'SOP':         { color: 'var(--task-color)',         hex: '#003b75', label: 'SOP' },
    'Certificate': { color: 'var(--task-color)',         hex: '#003b75', label: 'Cert' },
    'PPAP':        { color: 'var(--task-color)',         hex: '#003b75', label: 'PPAP' },
    'Top-level initiative': { color: 'var(--task-color)', hex: '#003b75', label: 'Init' },
};

const STANDARD_TASK_HEX = '#003b75';
const MILESTONE_HEX = '#f4a6a6';

const STATUS_COLORS = {
    'Done': '#48bb78',
    'Closed': '#48bb78',
    '完了': '#48bb78',
    'Done_ja': '#48bb78',
    'To Do': '#f6ad55',
    'In Progress': '#63b3ed',
    '騾ｲ陦御ｸｭ': '#63b3ed',
    '洶In Progress': '#63b3ed',
    'On Hold': '#f6ad55',
    '洫On Hold': '#f6ad55',
    'Start': '#ecc94b',
    'Process': '#68d391',
};

// Jira CSV column mapping requested for this standalone version.
// A=Summary, B=Issue Key, D=Issue Type, E=Status, X=End Date,
// DN=Grouping, EQ=Start Date, EM=Start Date (Fallback).
const JIRA_COLUMNS = {
    summary: 0,
    issueKey: 1,
    issueType: 3,
    status: 4,
    endDate: 23,
    grouping: 113,
    startDateFallback: 142,
    startDate: 146,
};

const SIMPLE_CSV_HEADERS = ['Summary', 'Issue Key', 'Grouping', 'Issue Type', 'Status', 'Start Date', 'End Date'];

function normalizeHeaderName(header) {
    return String(header || '')
        .replace(/^\ufeff/, '')
        .toLowerCase()
        .replace(/\s+/g, '');
}

// ========================================
// CSV Parser
// ========================================
function* parseCSVGenerator(text) {
    let current = '';
    let inQuotes = false;
    let fields = [];
    
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') {
            if (inQuotes && text[i + 1] === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (ch === ',' && !inQuotes) {
            fields.push(current.trim());
            current = '';
        } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
            if (ch === '\r' && text[i + 1] === '\n') i++;
            fields.push(current.trim());
            current = '';
            if (fields.some(f => f !== '')) {
                yield fields;
            }
            fields = [];
        } else {
            current += ch;
        }
    }
    if (current || fields.length > 0) {
        fields.push(current.trim());
        if (fields.some(f => f !== '')) {
            yield fields;
        }
    }
}

// ========================================
// Date Parser
// ========================================
function parseDate(dateStr) {
    if (!dateStr || dateStr.trim() === '') return null;
    
    let str = dateStr.trim();
    
    // Remove everything from the first HH:MM onwards.
    // e.g. "06/12/26 12:00 AM" -> "06/12/26"
    str = str.replace(/\s+\d{1,2}:\d{2}.*$/, '').trim();

    // Support simple CSV exports that include a weekday suffix.
    // e.g. "2027/8/31 Tue" or "2027/8/31(火)" -> "2027/8/31"
    str = str
        .replace(/\s+(sun|mon|tue|wed|thu|fri|sat)$/i, '')
        .replace(/[（(]\s*[日月火水木金土]\s*[）)]$/, '')
        .trim();
    
    // Format: DD/M/YY  (Jira export: day first, 2-digit year last  e.g. "06/12/26" = Dec 6 2026)
    const slashMatch = str.match(/^(\d{1,4})\/(\d{1,2})\/(\d{1,4})$/);
    if (slashMatch) {
        const a = parseInt(slashMatch[1]);
        const b = parseInt(slashMatch[2]);
        const c = parseInt(slashMatch[3]);
        let year, month, day;
        if (a > 31) {
            // YYYY/M/D  (4-digit year first)
            year = a; month = b - 1; day = c;
        } else if (c <= 99) {
            // DD/M/YY  (2-digit year last) 遯ｶ繝ｻJira format
            day = a; month = b - 1; year = c + 2000;
        } else {
            // D/M/YYYY (4-digit year last)
            day = a; month = b - 1; year = c;
        }
        const result = new Date(year, month, day);
        if (!isNaN(result.getTime())) return result;
        return null;
    }
    
    // Format: YYYY-MM-DD
    const dashMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (dashMatch) {
        return new Date(parseInt(dashMatch[1]), parseInt(dashMatch[2]) - 1, parseInt(dashMatch[3]));
    }
    
    // Try native parse as last resort
    const d = new Date(str);
    if (!isNaN(d.getTime())) return d;
    
    return null;
}

// Format a Date object to YYYY-MM-DD string for server API
function formatDate(date) {
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

// ========================================
// File Loading
// ========================================
function loadCSVFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const text = e.target.result;
            const tasks = parseCSVToTasks(text, file.name);
            resolve(tasks);
        };
        reader.onerror = reject;
        // Try UTF-8 first
        reader.readAsText(file, 'UTF-8');
    });
}

function parseCSVToTasks(text, filename) {
    const tasks = [];
    let headers = null;
    const get = (row, index) => (row[index] || '').trim();
    let colMap = { ...JIRA_COLUMNS };
    
    for (const row of parseCSVGenerator(text)) {
        if (!headers) {
            headers = row.map(normalizeHeaderName);

            const findHeaderIndex = (...matchers) => {
                return headers.findIndex(header => matchers.some(matcher => (
                    typeof matcher === 'string' ? header === matcher : matcher.test(header)
                )));
            };

            const summaryIdx = findHeaderIndex('summary');
            if (summaryIdx !== -1) colMap.summary = summaryIdx;

            const issueKeyIdx = findHeaderIndex('issuekey');
            if (issueKeyIdx !== -1) colMap.issueKey = issueKeyIdx;

            const issueTypeIdx = findHeaderIndex('issuetype');
            if (issueTypeIdx !== -1) colMap.issueType = issueTypeIdx;

            const statusIdx = findHeaderIndex('status');
            if (statusIdx !== -1) colMap.status = statusIdx;
            
            // Dynamically find critical custom columns since they can shift alphabetically
            const grpIdx = headers.findIndex(h => h.includes('grouping'));
            if (grpIdx !== -1) colMap.grouping = grpIdx;
            
            let startIdx = headers.findIndex(h => h.includes('startdate') || h === '開始日');
            if (startIdx === -1) {
                startIdx = headers.findIndex(h => h.includes('targetstart'));
            }
            if (startIdx !== -1) colMap.startDate = startIdx;
            
            let endIdx = headers.findIndex(h => h === '期限' || h === 'duedate');
            if (endIdx === -1) {
                endIdx = headers.findIndex(h => h.includes('targetend') || h.includes('actualend'));
            }
            if (endIdx === -1) {
                endIdx = headers.findIndex(h => h.includes('enddate'));
            }
            if (endIdx !== -1) colMap.endDate = endIdx;
            
            continue;
        }
        
        // Current Jira input mapping:
        // DN=Grouping, D=Issue Type, A=Summary, ET=Start Date, X=End Date, E=Status.
        // B=Issue Key is kept as the internal unique key.
        const key      = get(row, colMap.issueKey) || `${filename}-${tasks.length + 1}`;
        const grouping = get(row, colMap.grouping);
        const type     = get(row, colMap.issueType) || 'Task';
        const summary  = get(row, colMap.summary) || key;
        const status   = get(row, colMap.status) || 'To Do';
        const startDate = get(row, colMap.startDate) || get(row, colMap.startDateFallback);
        const deadline  = get(row, colMap.endDate);
        
        // Determine effective start and end dates
        const effectiveStart = parseDate(startDate);
        const effectiveEnd = parseDate(deadline);
        
        const typeLower = type.toLowerCase();
        const summaryUpper = String(summary).toUpperCase();
        const isMilestone = typeLower.includes('milestone') || typeLower.includes('sop') || summaryUpper.includes('SOP') || summaryUpper.includes('OEM LO');
        
        if (key && (effectiveStart || effectiveEnd)) {
            tasks.push({
                key,
                displayKey: isMilestone ? 'Milestone' : (grouping || ''),
                type,
                parent: '',
                summary,
                status,
                assignee: grouping || '',
                grouping: isMilestone ? 'Milestone' : grouping,
                startDate: effectiveStart,
                endDate: effectiveEnd,
                startDateRaw: startDate,
                endDateRaw: deadline,
                source: filename,
                project: grouping || key.split('-')[0],
                sourceOrder: tasks.length,
            });
        }
    }
    
    return tasks;
}

function escapeCSVCell(value) {
    const text = value == null ? '' : String(value);
    if (/[",\r\n]/.test(text)) {
        return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
}

function buildSimpleCSV(tasks) {
    const rows = [SIMPLE_CSV_HEADERS];
    tasks
        .filter(task => !task.isGroup)
        .forEach(task => {
            rows.push([
                task.summary || '',
                task.key || '',
                task.grouping || task.displayKey || '',
                task.type || 'Task',
                task.status || '',
                formatDate(task.startDate),
                formatDate(task.endDate),
            ]);
        });

    return rows
        .map(row => row.map(escapeCSVCell).join(','))
        .join('\r\n');
}

function getSimpleCsvFilename() {
    const stamp = new Date()
        .toISOString()
        .replace(/[-:T]/g, '')
        .slice(0, 12);
    return `GanttSimple_${stamp}.csv`;
}

function downloadSimpleCSV() {
    const tasks = state.tasks.filter(task => !task.isGroup);
    if (!tasks.length) {
        showNotification('No tasks available for simple CSV.', 'error');
        return;
    }

    const csv = '\ufeff' + buildSimpleCSV(tasks);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = getSimpleCsvFilename();
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function updateSimpleCsvDownloadButton() {
    const btn = document.getElementById('btn-download-simple-csv');
    if (!btn) return;
    btn.classList.toggle('hidden', state.tasks.filter(task => !task.isGroup).length === 0);
}

// ========================================
// Task Filtering & Grouping
// ========================================
function filterAndGroupTasks() {
    let tasks = [...state.tasks];
    
    // Filter by project
    if (state.selectedProject !== 'all') {
        tasks = tasks.filter(t => t.project === state.selectedProject);
    }
    
    // Only show tasks with dates
    const withDates = tasks.filter(t => t.startDate || t.endDate);
    const sortByDateThenSource = (a, b) => {
        const da = a.startDate || a.endDate || new Date(9999, 0);
        const db = b.startDate || b.endDate || new Date(9999, 0);
        if (da - db !== 0) return da - db;
        return (a.sourceOrder || 0) - (b.sourceOrder || 0);
    };

    const sortByGroupingAndKey = (a, b) => {
        const ga = getGroupingNumber(a);
        const gb = getGroupingNumber(b);
        if (ga !== gb) return ga - gb;
        
        if (isStepMeetingTask(a) && isStepMeetingTask(b)) {
            const da = a.startDate || a.endDate || new Date(9999, 0);
            const db = b.startDate || b.endDate || new Date(9999, 0);
            if (da - db !== 0) return da - db;
        }

        const keyCompare = compareIssueKey(a.key, b.key);
        if (keyCompare !== 0) return keyCompare;
        return sortByDateThenSource(a, b);
    };

    const sortByGroupingAndDate = (a, b) => {
        const ga = getGroupingNumber(a);
        const gb = getGroupingNumber(b);
        if (ga !== gb) return ga - gb;
        
        if (isStepMeetingTask(a) && isStepMeetingTask(b)) {
            const da = a.startDate || a.endDate || new Date(9999, 0);
            const db = b.startDate || b.endDate || new Date(9999, 0);
            if (da - db !== 0) return da - db;
        }
        
        const da = a.startDate || a.endDate || new Date(9999, 0);
        const db = b.startDate || b.endDate || new Date(9999, 0);
        if (da - db !== 0) return da - db;
        
        const keyCompare = compareIssueKey(a.key, b.key);
        if (keyCompare !== 0) return keyCompare;
        return sortByDateThenSource(a, b);
    };

    const sortedTasks = withDates.sort(sortByGroupingAndDate);
    
    state.filteredTasks = sortedTasks;
    return sortedTasks;
}

function getDisplaySummary(summary) {
    if (!summary) return '';
    const clean = summary.trim();
    if (clean.toLowerCase().includes('condition check')) {
        return 'CPK';
    }
    return clean;
}

function getDisplayGroupingName(groupName) {
    const text = String(groupName || '').trim();
    return text.replace(/^\s*\d+(?:\.\d+)?\.?\s*/, '') || text;
}

function isMilestoneTask(task) {
    return task.grouping === 'Milestone' || String(task.type || '').trim().toLowerCase().includes('milestone');
}

function isStepMeetingTask(task) {
    return String(task.grouping || '').trim().toLowerCase().includes('step meeting');
}

function getGroupingNumber(task) {
    const grouping = String(task.grouping || task.displayKey || task.project || '').trim();
    const gLower = grouping.toLowerCase();
    if (gLower.includes('milestone')) return 1;
    if (gLower.includes('step meeting')) return 2; // Step meeting placed directly after Milestone, before 04. Design
    
    const match = grouping.match(/^(\d+)/);
    return match ? parseInt(match[1], 10) : 9999;
}

const GROUP_STYLE = {
    'milestone': {
        bgBand: 'EBF3FC',        // very light blue row background
        groupBlock: 'DDEFFF',    // light blue block background
        groupText: '1F2937',     // dark gray
        summaryBg: 'E5E7EB',     // light gray
        summaryText: '1F2937',   // dark text
        barColor1: '8A1F1F',     // dark red for star
        barColor2: '8A1F1F'
    },
    'step meeting': {
        bgBand: 'FFFFFF',        // white background band
        groupBlock: '00B0F0',    // sky blue
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',     // light gray
        summaryText: '1F2937',   // dark text
        barColor1: '0050A0',     // PPT gradient start
        barColor2: '002060'      // PPT gradient end
    },
    '04. design': {
        bgBand: 'F2F4F7',        // gray background band
        groupBlock: '00B0F0',    // sky blue
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    },
    '08. investment': {
        bgBand: 'FFFFFF',        // white background band
        groupBlock: '00B0F0',
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    },
    '09. equipment': {
        bgBand: 'F2F4F7',        // gray background band
        groupBlock: '00B0F0',
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    },
    '11. validation': {
        bgBand: 'FFFFFF',        // white background band
        groupBlock: '00B0F0',
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    },
    '12. production': {
        bgBand: 'F2F4F7',        // gray background band
        groupBlock: '00B0F0',
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    },
    '13. closure': {
        bgBand: 'FFFFFF',        // white background band
        groupBlock: '00B0F0',
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    },
    'default': {
        bgBand: 'F2F4F7',        // gray background band
        groupBlock: '00B0F0',
        groupText: 'FFFFFF',
        summaryBg: 'E5E7EB',
        summaryText: '1F2937',
        barColor1: '0050A0',
        barColor2: '002060'
    }
};

function getGroupStyle(groupName) {
    const name = String(groupName || '').toLowerCase();
    if (name.includes('milestone')) return GROUP_STYLE['milestone'];
    if (name.includes('step meeting')) return GROUP_STYLE['step meeting'];
    if (name.includes('04') || name.includes('design')) return GROUP_STYLE['04. design'];
    if (name.includes('08') || name.includes('investment')) return GROUP_STYLE['08. investment'];
    if (name.includes('09') || name.includes('equipment')) return GROUP_STYLE['09. equipment'];
    if (name.includes('11')) return GROUP_STYLE['11. validation'];
    if (name.includes('12')) return GROUP_STYLE['12. production'];
    if (name.includes('13')) return GROUP_STYLE['13. closure'];
    return GROUP_STYLE['default'];
}

function getGroupingBackground(task) {
    const groupName = task.displayKey || task.grouping || '';
    const style = getGroupStyle(groupName);
    return '#' + style.bgBand;
}

function getGroupingPptFill(task) {
    const num = getGroupingNumber(task);
    if (num === 9999) return 'FFFFFF';
    return (num % 2 === 1) ? 'FFFFFF' : 'EAEAEA';
}

function compareIssueKey(a, b) {
    const parse = key => {
        const text = String(key || '');
        const match = text.match(/^(.*?)(\d+)$/);
        return match
            ? { prefix: match[1], number: parseInt(match[2], 10), text }
            : { prefix: text, number: Number.MAX_SAFE_INTEGER, text };
    };
    const ka = parse(a);
    const kb = parse(b);
    const prefixCompare = ka.prefix.localeCompare(kb.prefix, undefined, { numeric: true, sensitivity: 'base' });
    if (prefixCompare !== 0) return prefixCompare;
    if (ka.number !== kb.number) return ka.number - kb.number;
    return ka.text.localeCompare(kb.text, undefined, { numeric: true, sensitivity: 'base' });
}

// ========================================
// Timeline Calculations
// ========================================
function getTimelineBounds() {
    const tasks = state.filteredTasks.filter(t => !t.isGroup);
    if (tasks.length === 0) return { start: new Date(), end: new Date() };
    
    let earliest = null;
    let latest = null;
    
    for (const task of tasks) {
        const s = task.startDate;
        const e = task.endDate;
        
        if (s && (!earliest || s < earliest)) earliest = s;
        if (e && (!latest || e > latest)) latest = e;
        if (s && (!latest || s > latest)) latest = s;
        if (e && (!earliest || e < earliest)) earliest = e;
    }
    
    if (!earliest) earliest = new Date();
    if (!latest) latest = new Date();
    
    // Use custom bounds if set by user, otherwise auto-detect with padding
    const start = state.customStart 
        ? new Date(state.customStart.getFullYear(), state.customStart.getMonth(), 1)
        : new Date(earliest.getFullYear(), earliest.getMonth() - 1, 1);
    const end = state.customEnd 
        ? new Date(state.customEnd.getFullYear(), state.customEnd.getMonth() + 1, 0)
        : new Date(latest.getFullYear(), latest.getMonth() + 2, 0);
    
    return { start, end };
}

function getMonthsBetween(start, end) {
    const months = [];
    const current = new Date(start.getFullYear(), start.getMonth(), 1);
    
    while (current <= end) {
        months.push(new Date(current));
        current.setMonth(current.getMonth() + 1);
    }
    
    return months;
}

function daysBetween(a, b) {
    return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

function daysInMonth(date) {
    return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

// ========================================
// Rendering
// ========================================
function renderGantt() {
    try {
        const tasks = filterAndGroupTasks();
        if (tasks.length === 0) {
            showNotification('No tasks with valid dates were found in the selected project.', 'error');
            document.getElementById('gantt-section').classList.add('hidden');
            document.getElementById('controls-section').classList.add('hidden');
            return;
        }
        
        const bounds = getTimelineBounds();
        const months = getMonthsBetween(bounds.start, bounds.end);
        
        // Adjust day width based on view mode
        switch (state.viewMode) {
            case 'month':  state.dayWidth = 4; break;
            case 'quarter': state.dayWidth = 2; break;
            case 'year':   state.dayWidth = 1; break;
        }
        
        const containerEl = document.getElementById('gantt-container');
        if (containerEl) {
            if (state.hideSummaryColumn) {
                containerEl.classList.add('hide-summary-col');
            } else {
                containerEl.classList.remove('hide-summary-col');
            }
        }
        
        renderLabels(tasks);
        renderTimelineHeader(months);
        renderTimelineBody(tasks, bounds, months);
        
        // Show sections
        document.getElementById('gantt-section').classList.remove('hidden');
        document.getElementById('controls-section').classList.remove('hidden');
        
        showNotification(`Successfully rendered ${tasks.length} tasks!`, 'success');
    } catch (err) {
        console.error('Render error:', err);
        showNotification('Error rendering Gantt chart: ' + err.message, 'error');
    }
}

function renderLabels(tasks) {
    const container = document.getElementById('gantt-labels');
    container.innerHTML = '';
    
    // Header
    const header = document.createElement('div');
    header.className = 'gantt-label-header';
    header.innerHTML = `
        <span style="min-width:30px"></span>
        <span style="min-width:95px">Timeline</span>
        ${state.hideSummaryColumn ? '' : '<span style="flex:1;min-width:190px">Summary</span>'}
        <span style="min-width:70px">Type</span>
        <span style="min-width:90px">Status</span>
        ${state.hideSummaryColumn ? '' : `
        <span style="min-width:85px">Start</span>
        <span style="min-width:85px">End</span>
        `}
    `;
    container.appendChild(header);
    
    // Rows
    tasks.forEach((item, index) => {
        const row = document.createElement('div');
        
        if (item.isGroup) {
            row.className = 'gantt-label-row group-header';
            row.innerHTML = `<span style="flex:1">${item.label} (${item.count})</span>`;
        } else {
            const isChild = item._isChild || item.parent;
            row.className = `gantt-label-row${isChild ? ' child-row' : ''}`;
            row.style.background = getGroupingBackground(item);
            
            const style = getGroupStyle(item.displayKey || item.grouping || '');
            const typeConf = TYPE_CONFIG[item.type] || { color: 'var(--default-color)', hex: '#828282', label: '?' };
            row.innerHTML = `
                <div class="move-controls">
                    <button class="move-btn" onclick="moveTask('${item.key}', -1)">^</button>
                    <button class="move-btn" onclick="moveTask('${item.key}', 1)">v</button>
                </div>
                <span class="task-key">${item.displayKey || item.key}</span>
                ${state.hideSummaryColumn ? '' : `<span class="task-name editable" contenteditable="true" style="background:#${style.summaryBg};color:#fff;border-radius:4px;padding:3px 8px;text-align:center;font-weight:600" onblur="updateTask('${item.key}', 'summary', this.innerText)">${item.summary}</span>`}
                <span class="task-type-badge" style="background:${typeConf.hex}22;color:${typeConf.hex}">${typeConf.label}</span>
                <span class="task-status">${item.status || ''}</span>
                ${state.hideSummaryColumn ? '' : `
                <span class="task-date editable" contenteditable="true" onblur="updateTask('${item.key}', 'startDate', this.innerText)">${formatDate(item.startDate)}</span>
                <span class="task-date editable" contenteditable="true" onblur="updateTask('${item.key}', 'endDate', this.innerText)">${formatDate(item.endDate)}</span>
                `}
            `;
        }
        
        container.appendChild(row);
    });
}

function moveTask(key, direction) {
    const idx = state.tasks.findIndex(t => t.key === key);
    if (idx === -1) return;
    const newIdx = idx + direction;
    if (newIdx < 0 || newIdx >= state.tasks.length) return;
    
    const temp = state.tasks[idx];
    state.tasks[idx] = state.tasks[newIdx];
    state.tasks[newIdx] = temp;
    
    renderGantt();
}

function updateTask(key, field, value) {
    const task = state.tasks.find(t => t.key === key);
    if (!task) return;
    
    if (field === 'startDate' || field === 'endDate') {
        const d = parseDate(value);
        if (d) {
            task[field] = d;
            renderGantt(); // Re-render to update bar position
        } else {
            // Restore if invalid
            renderGantt();
        }
    } else {
        task[field] = value;
        // Don't necessarily need to re-render everything for summary edit unless bar has label
        renderGantt();
    }
}

function renderTimelineHeader(months) {
    const container = document.getElementById('gantt-timeline-header');
    container.innerHTML = '';
    
    // Group months by year
    const years = {};
    months.forEach(m => {
        const y = m.getFullYear();
        if (!years[y]) years[y] = [];
        years[y].push(m);
    });

    // Create Year Row
    const yearRow = document.createElement('div');
    yearRow.className = 'timeline-year-row';
    for (const year in years) {
        const yearMonths = years[year];
        const width = yearMonths.reduce((sum, m) => sum + daysInMonth(m) * state.dayWidth, 0);
        const yearEl = document.createElement('div');
        yearEl.className = 'timeline-year-label';
        yearEl.style.width = `${width}px`;
        yearEl.textContent = year;
        yearRow.appendChild(yearEl);
    }
    container.appendChild(yearRow);

    // Create Month Row
    const monthRow = document.createElement('div');
    monthRow.className = 'timeline-month-row';
    for (const month of months) {
        const days = daysInMonth(month);
        const width = days * state.dayWidth;
        const monthEl = document.createElement('div');
        monthEl.className = 'timeline-month-label-hier';
        monthEl.style.width = `${width}px`;
        monthEl.textContent = month.getMonth() + 1;
        monthRow.appendChild(monthEl);
    }
    container.appendChild(monthRow);
}

function attachTimelineLabelEditor(label, item) {
    label.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        const newSummary = prompt('Edit Task Name / 繧ｿ繧ｹ繧ｯ蜷阪ｒ邱ｨ髮・', item.summary);
        if (newSummary !== null && newSummary.trim() !== '') {
            updateTask(item.key, 'summary', newSummary.trim());
        }
    });
}

function placeStepMeetingLabelBefore(rowEl, displaySummary, startOffset, item) {
    const label = document.createElement('div');
    label.className = 'gantt-timeline-label outside step-meeting-before';
    label.textContent = displaySummary;
    label.style.left = `${Math.max(0, startOffset - 204)}px`;
    label.style.width = `${Math.max(96, Math.min(160, startOffset - 44))}px`;
    label.style.textAlign = 'right';
    label.style.overflow = 'hidden';
    label.style.textOverflow = 'ellipsis';

    attachTimelineLabelEditor(label, item);
    rowEl.appendChild(label);
}

function renderTimelineBody(tasks, bounds, months) {
    const container = document.getElementById('gantt-timeline-body');
    container.innerHTML = '';
    
    // Calculate total width
    let totalWidth = 0;
    for (const month of months) {
        totalWidth += daysInMonth(month) * state.dayWidth;
    }
    container.style.width = `${totalWidth}px`;
    
    // Draw grid lines (month boundaries)
    let xOffset = 0;
    for (const month of months) {
        const line = document.createElement('div');
        line.className = 'gantt-grid-line month-line';
        line.style.left = `${xOffset}px`;
        container.appendChild(line);
        xOffset += daysInMonth(month) * state.dayWidth;
    }
    
    const today = new Date();
    const todayOffset = daysBetween(bounds.start, today) * state.dayWidth;
    if (todayOffset >= 0 && todayOffset <= totalWidth) {
        const todayLine = document.createElement('div');
        todayLine.className = 'gantt-today-line';
        todayLine.style.left = `${todayOffset}px`;
        container.appendChild(todayLine);
    }
    
    // Task bars
    for (let i = 0; i < tasks.length; i++) {
        const item = tasks[i];
        const rowEl = document.createElement('div');
        rowEl.className = `gantt-row${item.isGroup ? ' group-header' : ''}`;
        if (!item.isGroup) rowEl.style.background = getGroupingBackground(item);
        
        if (!item.isGroup && (item.startDate || item.endDate)) {
            const start = item.startDate || item.endDate;
            const end = item.endDate || item.startDate;
            const rangeStart = start < end ? start : end;
            const rangeEnd = end >= start ? end : start;

            if (rangeEnd < bounds.start || rangeStart > bounds.end) {
                container.appendChild(rowEl);
                continue;
            }

            const visibleStart = rangeStart < bounds.start ? bounds.start : rangeStart;
            const visibleEnd = rangeEnd > bounds.end ? bounds.end : rangeEnd;
            
            const startOffset = daysBetween(bounds.start, visibleStart) * state.dayWidth;
            const duration = Math.max(daysBetween(visibleStart, visibleEnd), 1) * state.dayWidth;
            
            const style = getGroupStyle(item.displayKey || item.grouping || '');
            const isMilestone = isMilestoneTask(item);
            const isStepMeeting = isStepMeetingTask(item);
            
            const bar = document.createElement('div');
            bar.className = `gantt-bar${isMilestone ? ' milestone' : isStepMeeting ? ' diamond' : ''}`;
            bar.style.left = `${startOffset}px`;
            
            if (isMilestone) {
                bar.style.background = 'transparent';
                bar.style.color = '#' + style.barColor1;
                bar.textContent = '*';
            } else if (isStepMeeting) {
                bar.style.background = 'transparent';
                bar.style.color = '#' + style.barColor1;
                bar.textContent = '◆';
            } else {
                bar.style.width = `${Math.max(duration, 6)}px`;
                bar.style.background = '#' + style.barColor2;
            }
            
            // Tooltip data
            bar.dataset.key = item.key;
            bar.dataset.summary = item.summary;
            bar.dataset.type = item.type;
            bar.dataset.status = item.status;
            bar.dataset.start = formatDate(item.startDate);
            bar.dataset.end = formatDate(item.endDate);
            bar.dataset.assignee = item.assignee;
            
            bar.addEventListener('mouseenter', showTooltip);
            bar.addEventListener('mouseleave', hideTooltip);
            bar.addEventListener('mousemove', moveTooltip);
            
            // Stagger animation
            bar.style.animationDelay = `${i * 0.02}s`;
            
            rowEl.appendChild(bar);

            // Add text label and date annotations inside timeline
            if (state.hideSummaryColumn) {
                const displaySummary = getDisplaySummary(item.summary);
                const isInfPpap = displaySummary.toLowerCase().includes('inf ppap');

                // Short Date Formatting: "M/D"
                const formatShortDate = (d) => d ? `${d.getMonth() + 1}/${d.getDate()}` : '';
                const startStr = formatShortDate(item.startDate || item.endDate);
                const endStr = formatShortDate(item.endDate || item.startDate);
                
                // 1. Start Date Label (printed to the left of the bar/icon)
                const startLabel = document.createElement('div');
                startLabel.className = 'gantt-timeline-date start';
                startLabel.textContent = startStr;
                startLabel.style.left = `${startOffset - 36}px`;
                startLabel.style.width = '30px';
                startLabel.style.textAlign = 'right';
                rowEl.appendChild(startLabel);

                if (isMilestone) {
                    // For milestone/step meetings, there's only one date, so we only print start date.
                    // The summary label is placed to the right of the icon.
                    const label = document.createElement('div');
                    label.className = 'gantt-timeline-label outside';
                    label.textContent = displaySummary;
                    label.style.left = `${startOffset + 24}px`;
                    
                    label.addEventListener('dblclick', (e) => {
                        e.stopPropagation();
                        const newSummary = prompt('Edit Task Name / タスク名を編集:', item.summary);
                        if (newSummary !== null && newSummary.trim() !== '') {
                            updateTask(item.key, 'summary', newSummary.trim());
                        }
                    });
                    
                    rowEl.appendChild(label);
                } else if (isStepMeeting) {
                    placeStepMeetingLabelBefore(rowEl, displaySummary, startOffset, item);
                } else {
                    // 2. End Date Label
                    const endLabel = document.createElement('div');
                    endLabel.className = 'gantt-timeline-date end';
                    endLabel.textContent = endStr;
                    endLabel.style.left = `${startOffset + duration + 6}px`;
                    rowEl.appendChild(endLabel);

                    // 3. Summary Label placement (try inside, shrink font, or place outside)
                    const fontSizes = [11, 10, 9, 8, 7];
                    const charWidths = { 11: 6.5, 10: 6.0, 9: 5.5, 8: 4.8, 7: 4.2 };
                    
                    let fits = false;
                    let chosenFontSize = 11;
                    if (!isInfPpap) {
                        for (const fs of fontSizes) {
                            const estTextWidth = displaySummary.length * charWidths[fs];
                            if (duration >= estTextWidth + 12) {
                                chosenFontSize = fs;
                                fits = true;
                                break;
                            }
                        }
                    }
                    
                    if (fits) {
                        // Place inside the bar!
                        bar.textContent = displaySummary;
                        bar.style.fontSize = `${chosenFontSize}px`;
                        bar.style.justifyContent = 'center';
                        bar.style.textAlign = 'center';
                        bar.style.color = '#FFFFFF';
                        bar.style.fontWeight = '600';
                        bar.style.padding = '0 4px';
                        bar.style.whiteSpace = 'nowrap';
                        bar.style.overflow = 'hidden';
                        bar.style.textOverflow = 'ellipsis';
                        
                        // Make it editable on double click as well
                        bar.addEventListener('dblclick', (e) => {
                            e.stopPropagation();
                            const newSummary = prompt('Edit Task Name / タスク名を編集:', item.summary);
                            if (newSummary !== null && newSummary.trim() !== '') {
                                updateTask(item.key, 'summary', newSummary.trim());
                            }
                        });
                    } else {
                        // Place outside (on the right), starting after the End Date Label
                        const label = document.createElement('div');
                        label.className = 'gantt-timeline-label outside';
                        label.textContent = (isInfPpap ? '← ' : '') + displaySummary;
                        label.style.left = `${startOffset + duration + 38}px`;
                        
                        label.addEventListener('dblclick', (e) => {
                            e.stopPropagation();
                            const newSummary = prompt('Edit Task Name / タスク名を編集:', item.summary);
                            if (newSummary !== null && newSummary.trim() !== '') {
                                updateTask(item.key, 'summary', newSummary.trim());
                            }
                        });
                        
                        rowEl.appendChild(label);
                    }
                }
            }
        }
        
        container.appendChild(rowEl);
    }
}

// ========================================
// Tooltip
// ========================================
let tooltipEl = null;

function createTooltip() {
    tooltipEl = document.createElement('div');
    tooltipEl.className = 'gantt-tooltip';
    tooltipEl.id = 'gantt-tooltip';
    document.body.appendChild(tooltipEl);
}

function showTooltip(e) {
    if (!tooltipEl) createTooltip();
    
    const bar = e.currentTarget;
    tooltipEl.innerHTML = `
        <div class="tooltip-title">${bar.dataset.summary}</div>
        <div class="tooltip-row"><span>Key:</span><span>${bar.dataset.key}</span></div>
        <div class="tooltip-row"><span>Type:</span><span>${bar.dataset.type}</span></div>
        <div class="tooltip-row"><span>Status:</span><span>${bar.dataset.status}</span></div>
        <div class="tooltip-row"><span>Start:</span><span>${bar.dataset.start || 'Not set'}</span></div>
        <div class="tooltip-row"><span>End:</span><span>${bar.dataset.end || 'Not set'}</span></div>
        ${bar.dataset.assignee ? `<div class="tooltip-row"><span>Assignee:</span><span>${bar.dataset.assignee}</span></div>` : ''}
    `;
    tooltipEl.classList.add('visible');
}

function hideTooltip() {
    if (tooltipEl) tooltipEl.classList.remove('visible');
}

function moveTooltip(e) {
    if (!tooltipEl) return;
    const x = e.clientX + 12;
    const y = e.clientY + 12;
    tooltipEl.style.left = `${x}px`;
    tooltipEl.style.top = `${y}px`;
}

// ========================================
// Date Formatting
// ========================================
function formatDate(date) {
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
}

function isLocalServer() {
    return ['localhost', '127.0.0.1', '::1'].includes(window.location.hostname);
}

function loadScript(src) {
    return new Promise((resolve, reject) => {
        const existing = document.querySelector(`script[src="${src}"]`);
        if (existing) {
            if (existing.dataset.loaded === 'true') {
                resolve();
                return;
            }
            if (existing.dataset.loading === 'true') {
                existing.addEventListener('load', resolve, { once: true });
                existing.addEventListener('error', () => reject(new Error(`Failed to load ${src}`)), { once: true });
                return;
            }
            resolve();
            return;
        }

        const script = document.createElement('script');
        script.src = src;
        script.dataset.loading = 'true';
        script.onload = () => {
            script.dataset.loaded = 'true';
            script.dataset.loading = 'false';
            resolve();
        };
        script.onerror = () => reject(new Error(`Failed to load ${src}`));
        document.head.appendChild(script);
    });
}

async function getPptxConstructor() {
    if (window.PptxGenJS || window.pptxgen || window.pptxgenjs) {
        return window.PptxGenJS || window.pptxgen || window.pptxgenjs;
    }

    const sourceSets = [
        ['vendor/jszip.min.js', 'vendor/pptxgen.min.js'],
        ['vendor/pptxgen.bundle.js'],
    ];

    for (const sources of sourceSets) {
        try {
            for (const src of sources) {
                await loadScript(src);
            }
            const constructor = window.PptxGenJS || window.pptxgen || window.pptxgenjs;
            if (constructor) return constructor;
        } catch (err) {
            console.warn(err.message);
        }
    }

    throw new Error('PptxGenJS library could not be loaded.');
}

async function captureGanttImage() {
    const ganttWrapper = document.getElementById('gantt-wrapper');
    const container = document.getElementById('gantt-container');
    const origWrapperOverflow = ganttWrapper.style.overflow;
    const origContainerOverflow = container.style.overflow;
    const origWrapperWidth = ganttWrapper.style.width;

    ganttWrapper.style.overflow = 'visible';
    container.style.overflow = 'visible';
    ganttWrapper.style.width = `${container.scrollWidth}px`;

    try {
        const canvas = await html2canvas(ganttWrapper, {
            backgroundColor: getComputedStyle(document.documentElement).getPropertyValue('--bg-primary').trim(),
            scale: 2,
            logging: false,
            useCORS: true,
            windowWidth: Math.max(document.documentElement.clientWidth, container.scrollWidth),
        });
        return canvas.toDataURL('image/png');
    } finally {
        ganttWrapper.style.overflow = origWrapperOverflow;
        container.style.overflow = origContainerOverflow;
        ganttWrapper.style.width = origWrapperWidth;
    }
}

const PPT_COLORS = {
    white: 'FFFFFF',
    black: '000000',
    bgGray: 'F8F9FB',
    grid: 'EEEEEE',
    monthGridLine: 'D6DEE8',
    navy: '1F3D6E',
    timelineTeal: '1F3D6E',
    timelineTask: 'BFE7F2',
    timelineMilestone: '8A1F1F',
    timelineLine: '274060',
    timelineToday: '7F8792',
    title: '1F2937',
    muted: '6B7280',
    lightBlue: 'BDD7EE',
    today: 'EF4444',
    milestoneText: 'C00000',
    defaultBar: '003B75',
};

const PPT_FONT_RULES = {
    milestoneLabel: 11,
    milestoneDate: 10.5,
    stepMeetingLabel: 11.5,
    stepMeetingDate: 10,
    trailingDate: 8,
    milestoneDateWidth: 0.52,
    stepMeetingDateWidth: 0.48,
    trailingDateWidth: 0.44,
};

const PPT_TYPE_COLORS = {
    'Milestone': 'F4A6A6',
    'Goal': '003B75',
    'Review': '003B75',
    'Event': '003B75',
    'Task': '003B75',
    'Sub task': '003B75',
    'Validation': '003B75',
    'SOP': '003B75',
    'PPAP': '003B75',
    'Certificate': '003B75',
    'Top-level initiative': '003B75',
};

function hexToRgb(hex) {
    const clean = hex.replace('#', '');
    return [
        parseInt(clean.slice(0, 2), 16),
        parseInt(clean.slice(2, 4), 16),
        parseInt(clean.slice(4, 6), 16),
    ];
}

function mixHex(hex, weight = 0.2) {
    const [r, g, b] = hexToRgb(hex);
    const mix = v => Math.min(255, Math.round(v * weight + 255 * (1 - weight)));
    return [mix(r), mix(g), mix(b)].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

function monthStart(date) {
    return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date, count) {
    return new Date(date.getFullYear(), date.getMonth() + count, 1);
}

function daysBetween(start, end) {
    return Math.round((end - start) / 86400000);
}

function daysInMonth(date) {
    return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

function pptDate(date) {
    if (!date) return '';
    return `${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
}

function pptText(slide, text, x, y, w, h, opts = {}) {
    slide.addText(String(text || ''), {
        x, y, w, h,
        fontFace: 'Meiryo',
        fontSize: opts.fontSize || 9,
        color: opts.color || PPT_COLORS.black,
        bold: Boolean(opts.bold),
        italic: Boolean(opts.italic),
        align: opts.align || 'left',
        valign: 'mid',
        margin: opts.margin ?? 0.03,
        fit: 'shrink',
        breakLine: false,
    });
}

function pptRect(slide, shapeType, x, y, w, h, fill, line = null) {
    let fillOption = { transparency: 100 };
    if (fill) {
        if (typeof fill === 'object') {
            fillOption = fill;
        } else {
            fillOption = { color: fill };
        }
    }
    
    let defaultLineColor = 'none';
    if (fill) {
        if (typeof fill === 'string') {
            defaultLineColor = fill;
        } else if (fill.color1) {
            defaultLineColor = fill.color1;
        }
    }

    slide.addShape(shapeType, {
        x, y, w: Math.max(w, 0.001), h: Math.max(h, 0.001),
        fill: fillOption,
        line: line || (defaultLineColor !== 'none' ? { color: defaultLineColor } : { color: 'none' }),
    });
}

function drawPptMonthGridLines(slide, shapes, months, chartL, chartW, totalDays, topY, bottomY) {
    const lineH = Math.max(bottomY - topY, 0.001);
    let elapsedDays = 0;

    const addMonthLine = (x) => {
        slide.addShape(shapes.line, {
            x,
            y: topY,
            w: 0,
            h: lineH,
            line: {
                color: PPT_COLORS.monthGridLine,
                width: 0.25,
                transparency: 12,
            },
        });
    };

    months.forEach(month => {
        const x = chartL + chartW * (elapsedDays / totalDays);
        addMonthLine(x);
        elapsedDays += daysInMonth(month);
    });

    addMonthLine(chartL + chartW);
}

function pptStepMeetingLabelBefore(slide, displaySummary, bx, y, h, fontSize, leftM) {
    // Step meeting label before the diamond, leaving space for the date beside the icon.
    const labelX = Math.max(leftM, bx - 1.95);
    const labelW = Math.max(0.22, bx - 0.72 - labelX);
    pptText(slide, displaySummary, labelX, y, labelW, h, {
        fontSize,
        color: '1F2937',
        align: 'right',
    });
}

function drawPptTodayLine(slide, shapes, startBound, endBound, chartL, chartW, totalDays, topY, height) {
    const today = new Date();
    if (today < startBound || today > endBound) return;

    const tx = chartL + chartW * (daysBetween(startBound, today) / totalDays);
    pptRect(slide, shapes.rect, tx, topY, 0.014, height, PPT_COLORS.today, { color: PPT_COLORS.today, width: 0 });
    pptText(slide, 'Today', tx - 0.18, topY - 0.16, 0.42, 0.15, {
        fontSize: 6.5,
        color: PPT_COLORS.today,
        bold: true,
        align: 'center',
        margin: 0,
    });
}

async function generatePPTXInBrowser(tasks, title, filename, options = {}) {
    const isFourThree = options.layout === '4:3';
    const hideSummary = options.hasOwnProperty('hideSummary') ? options.hideSummary : state.hideSummaryColumn;
    const PptxConstructor = await getPptxConstructor();
    const pptx = new PptxConstructor();
    if (isFourThree) {
        pptx.defineLayout({ name: 'DAICEL_4_3', width: 10, height: 7.5 });
        pptx.layout = 'DAICEL_4_3';
    } else {
        pptx.layout = 'LAYOUT_WIDE';
    }
    pptx.author = 'Gantt Chart Generator';
    pptx.subject = 'CSV Gantt export';
    pptx.title = title;
    pptx.company = 'Browser export';

    const shapes = pptx.ShapeType || {
        rect: 'rect',
        roundRect: 'roundRect',
        diamond: 'diamond',
        ellipse: 'ellipse',
        line: 'line',
    };

    const rows = tasks
        .map(t => {
            const start = t.startDate || t.endDate;
            const end = t.endDate || t.startDate;
            if (!start || !end) return null;
            return { ...t, start, end };
        })
        .filter(Boolean);

    if (!rows.length) {
        throw new Error('No tasks with valid dates found.');
    }

    const allDates = rows.flatMap(t => [t.start, t.end]);
    const startBound = state.customStart || new Date(Math.min(...allDates));
    const endBound = state.customEnd || new Date(Math.max(...allDates));
    const bs = monthStart(startBound);
    const be = state.customEnd ? addMonths(monthStart(endBound), 1) : addMonths(monthStart(endBound), 3);
    const totalDays = Math.max(1, daysBetween(bs, be));
    const months = [];
    for (let cur = new Date(bs); cur < be; cur = addMonths(cur, 1)) {
        months.push(new Date(cur));
    }

    const slideW = isFourThree ? 10 : 13.33;
    const slideH = isFourThree ? 7.5 : 7.5;
    const leftM = isFourThree ? 0.3 : 0.4;
    const topM = isFourThree ? 1.55 : 1.3;

    const maxSummaryLen = Math.max(10, ...rows.map(t => (t.summary || '').length));
    
    const wKey = isFourThree ? 1.35 : 1.55;
    const wSummary = hideSummary ? 0 : Math.min(3.5, Math.max(isFourThree ? 1.58 : 2.05, maxSummaryLen * 0.08));
    const wDate = isFourThree ? 0.50 : 0.58;
    const leftW = wKey + (hideSummary ? 0 : wSummary + wDate * 2);
    const chartL = leftM + leftW + (isFourThree ? 0.08 : 0.1);
    const chartW = slideW - chartL - (isFourThree ? 0.15 : 0.2);
    const hYear = isFourThree ? 0.22 : 0.25;
    const hMonth = isFourThree ? 0.22 : 0.25;
    const hHeader = hYear + hMonth;
    const hRow = hideSummary ? Math.min(0.32, (slideH - topM - 0.6) / rows.length) : (isFourThree ? 0.30 : 0.32);
    const barH = hRow * 0.7;
    const tasksPerSlide = hideSummary ? rows.length : (isFourThree ? 16 : 16);
    const pages = [];
    for (let i = 0; i < rows.length; i += tasksPerSlide) {
        pages.push(rows.slice(i, i + tasksPerSlide));
    }

    pages.forEach((page, pageIndex) => {
        const slide = pptx.addSlide();
        // Transparent slide background to allow PPT templates to show through

        if (isFourThree) {
            pptRect(slide, shapes.rect, 0, 0.02, slideW, 0.035, 'E60012');
            pptRect(slide, shapes.rect, 0, 0.075, slideW, 0.035, '00A3E0');
            pptRect(slide, shapes.rect, 0, 0.13, slideW, 0.035, '005BAC');
            pptText(slide, 'DAICEL', slideW - 1.35, 0.13, 1.1, 0.3, {
                fontSize: 18,
                bold: true,
                color: '00A3E0',
                align: 'right',
            });
            pptText(slide, String(pageIndex + 1), slideW / 2 - 0.15, slideH - 0.25, 0.3, 0.12, {
                fontSize: 8,
                color: PPT_COLORS.black,
                align: 'center',
                margin: 0,
            });
        }

        pptText(slide, title, leftM, isFourThree ? 0.35 : 0.1, isFourThree ? 5.6 : 6.0, 0.4, {
            fontSize: isFourThree ? 16 : 20,
            bold: true,
            color: PPT_COLORS.title,
        });
        pptText(slide, `${bs.getFullYear()}/${String(bs.getMonth() + 1).padStart(2, '0')} - ${be.getFullYear()}/${String(be.getMonth() + 1).padStart(2, '0')} (Page ${pageIndex + 1}/${pages.length})`,
            leftM, isFourThree ? 0.70 : 0.5, isFourThree ? 3.8 : 4.4, 0.2, { fontSize: isFourThree ? 8 : 9, color: PPT_COLORS.muted });

        let lx = leftM;
        const columns = [
            ['Timeline', wKey],
            ...(hideSummary ? [] : [
                ['Summary', wSummary],
                ['Start', wDate],
                ['End', wDate]
            ]),
        ];
        columns.forEach(([label, width]) => {
            pptRect(slide, shapes.rect, lx, topM - hHeader, width, hHeader, PPT_COLORS.navy, { color: PPT_COLORS.white, width: 0.5 });
            pptText(slide, label, lx, topM - hHeader, width, hHeader, {
                fontSize: isFourThree ? 9 : 9,
                bold: true,
                color: PPT_COLORS.white,
                align: 'center',
            });
            lx += width;
        });

        const years = [];
        months.forEach(m => {
            let group = years.find(y => y.year === m.getFullYear());
            if (!group) {
                group = { year: m.getFullYear(), months: [] };
                years.push(group);
            }
            group.months.push(m);
        });

        let yx = chartL;
        years.forEach(group => {
            const yw = group.months.reduce((sum, m) => sum + chartW * (daysInMonth(m) / totalDays), 0);
            pptRect(slide, shapes.rect, yx, topM - hHeader, yw, hYear, PPT_COLORS.timelineTeal, { color: PPT_COLORS.white, width: 0.5 });
            pptText(slide, group.year, yx, topM - hHeader, yw, hYear, {
                fontSize: isFourThree ? 9 : 10,
                bold: true,
                color: PPT_COLORS.white,
                align: 'center',
            });
            yx += yw;
        });

        let mx = chartL;
        months.forEach(m => {
            const mw = chartW * (daysInMonth(m) / totalDays);
            pptRect(slide, shapes.rect, mx, topM - hMonth, mw, hMonth, PPT_COLORS.lightBlue, { color: PPT_COLORS.white, width: 0.5 });
            pptText(slide, String(m.getMonth() + 1).padStart(2, '0'), mx, topM - hMonth, mw, hMonth, {
                fontSize: isFourThree ? 9 : 8,
                align: 'center',
            });
            mx += mw;
        });

        // 1. Calculate Grouping blocks first so we can draw shading under tasks
        const groupBlocks = [];
        for (let i = 0; i < page.length; ) {
            const currentGroup = page[i].displayKey || page[i].grouping || '';
            let count = 1;
            while (i + count < page.length) {
                const nextGroup = page[i + count].displayKey || page[i + count].grouping || '';
                if (nextGroup === currentGroup) count++;
                else break;
            }
            groupBlocks.push({ group: currentGroup, startIndex: i, count: count });
            i += count;
        }

        // 2. Draw block backgrounds (shading) for ALL groupings across the Start, End, and chart area
        groupBlocks.forEach((block) => {
            const by = topM + block.startIndex * hRow;
            const bh = block.count * hRow;
            const bgL = hideSummary ? chartL : leftM + wKey + wSummary;
            const bgW = hideSummary ? chartW : slideW - leftM - bgL;
            const style = getGroupStyle(block.group);
            pptRect(slide, shapes.rect, bgL, by, bgW, bh, style.bgBand);
        });

        drawPptMonthGridLines(slide, shapes, months, chartL, chartW, totalDays, topM, topM + page.length * hRow);

        // 3. Draw rows, cells, task bars, and labels
        page.forEach((row, rowIndex) => {
            const y = topM + rowIndex * hRow;
            const isMilestone = isMilestoneTask(row);
            const displaySummary = getDisplaySummary(row.summary);
            const isInfPpap = displaySummary.toLowerCase().includes('inf ppap');
            
            let curLx = leftM;
            curLx += wKey;

            const style = getGroupStyle(row.displayKey || row.grouping || '');

            // Summary cell, Start date cell, and End date cell (drawn only if !hideSummary)
            if (!hideSummary) {
                const boxX = curLx + 0.05;
                const boxY = y + 0.03;
                const boxW = wSummary - 0.1;
                const boxH = hRow - 0.06;
                pptRect(slide, shapes.roundRect, boxX, boxY, boxW, boxH, style.summaryBg, null);
                pptText(slide, displaySummary, boxX + 0.05, boxY, boxW - 0.1, boxH, { fontSize: isFourThree ? 9 : 9, color: style.summaryText, align: 'center' });
                curLx += wSummary;

                // Start cell: print date directly (no border/box)
                pptText(slide, pptDate(row.start), curLx, y, wDate, hRow, { fontSize: 9, align: 'center', color: '1F2937' });
                curLx += wDate;
                
                // End cell: print date directly (no border/box)
                pptText(slide, pptDate(row.end), curLx, y, wDate, hRow, { fontSize: 9, align: 'center', color: '1F2937' });
            }

            const rangeStart = row.start < row.end ? row.start : row.end;
            const rangeEnd = row.end >= row.start ? row.end : row.start;
            if (rangeEnd < bs || rangeStart > be) return;
            const visibleStart = rangeStart < bs ? bs : rangeStart;
            const visibleEnd = rangeEnd > be ? be : rangeEnd;
            const rawX = chartL + chartW * (daysBetween(bs, visibleStart) / totalDays);
            const rawW = chartW * (Math.max(1, daysBetween(visibleStart, visibleEnd)) / totalDays);
            const bx = Math.max(chartL, Math.min(chartL + chartW, rawX));
            const bw = Math.max(0.03, Math.min(rawW, chartL + chartW - bx));

            if (hideSummary) {
                const pptShortDate = d => d ? `${d.getMonth() + 1}/${d.getDate()}` : '';
                const startStr = pptShortDate(row.start);
                const endStr = pptShortDate(row.end);

                const currentBarH = hRow * 0.7;
                const barOffsetY = (hRow - currentBarH) / 2;
                if (isMilestoneTask(row)) {
                    // Milestone date label
                    pptText(slide, startStr, bx - PPT_FONT_RULES.milestoneDateWidth - 0.08, y + barOffsetY, PPT_FONT_RULES.milestoneDateWidth, currentBarH, {
                        fontSize: PPT_FONT_RULES.milestoneDate,
                        color: '6B7280',
                        align: 'right',
                    });
                    // Milestone star
                    pptText(slide, '*', bx - 0.04, y + barOffsetY - 0.02, 0.08, currentBarH + 0.04, {
                        fontSize: hRow < 0.20 ? 10 : 12,
                        bold: true,
                        color: PPT_COLORS.milestoneText,
                        align: 'center',
                        margin: 0,
                    });
                    // Milestone label to the right of the star
                    const labelX = bx + 0.06;
                    const labelW = slideW - leftM - labelX;
                    pptText(slide, displaySummary, labelX, y + barOffsetY, labelW, currentBarH, {
                        fontSize: PPT_FONT_RULES.milestoneLabel,
                        bold: true,
                        color: PPT_COLORS.milestoneText,
                        align: 'left',
                    });
                } else if (isStepMeetingTask(row)) {
                    pptText(slide, startStr, bx - PPT_FONT_RULES.stepMeetingDateWidth - 0.10, y + barOffsetY, PPT_FONT_RULES.stepMeetingDateWidth, currentBarH, {
                        fontSize: PPT_FONT_RULES.stepMeetingDate,
                        color: '6B7280',
                        align: 'right',
                    });
                    // Step meeting diamond shape
                    const diamondSize = Math.max(0.06, hRow - 0.10);
                    const diamondOffsetY = (hRow - diamondSize) / 2;
                    pptRect(slide, shapes.diamond, bx - diamondSize / 2, y + diamondOffsetY, diamondSize, diamondSize, style.barColor1);
                    pptStepMeetingLabelBefore(slide, displaySummary, bx, y + barOffsetY, currentBarH, PPT_FONT_RULES.stepMeetingLabel, leftM);
                } else {
                    // 1. Start Date label (printed to the left of the bar/icon)
                    pptText(slide, startStr, bx - PPT_FONT_RULES.trailingDateWidth - 0.05, y + barOffsetY, PPT_FONT_RULES.trailingDateWidth, currentBarH, {
                        fontSize: PPT_FONT_RULES.trailingDate,
                        color: '6B7280',
                        align: 'right',
                    });
                    // Task bar
                    pptRect(slide, shapes.rect, bx, y + barOffsetY, bw, currentBarH, style.barColor2);
                    
                    // 2. End Date label (printed to the right of the bar)
                    pptText(slide, endStr, bx + bw + 0.05, y + barOffsetY, PPT_FONT_RULES.trailingDateWidth, currentBarH, { fontSize: PPT_FONT_RULES.trailingDate, color: '6B7280', align: 'left' });

                    // 3. Smart label placement (inside vs outside)
                    const fontSizes = [8.5, 7.5, 6.5, 5.5];
                    const charWidths = { 8.5: 0.065, 7.5: 0.057, 6.5: 0.050, 5.5: 0.042 };
                    let fits = false;
                    let chosenFontSize = hRow < 0.20 ? 6.5 : (hRow < 0.25 ? 7.5 : 8.5);
                    if (!isInfPpap) {
                        for (const fs of fontSizes) {
                            const textW = displaySummary.length * charWidths[fs];
                            if (bw >= textW + 0.12) {
                                chosenFontSize = fs;
                                fits = true;
                                break;
                            }
                        }
                    }

                    if (fits) {
                        // Fit inside
                        pptText(slide, displaySummary, bx, y + barOffsetY, bw, currentBarH, {
                            fontSize: chosenFontSize,
                            color: 'FFFFFF',
                            bold: true,
                            align: 'center',
                        });
                    } else {
                        // Fit outside to the right of End Date
                        const labelText = (isInfPpap ? '<- ' : '') + displaySummary;
                        const labelX = bx + bw + 0.40;
                        const labelW = slideW - leftM - labelX;
                        pptText(slide, labelText, labelX, y + barOffsetY, labelW, currentBarH, {
                            fontSize: hRow < 0.20 ? 6.5 : (hRow < 0.25 ? 7.5 : 8.5),
                            color: '1F2937',
                            align: 'left',
                        });
                    }
                }
            } else {
                const currentBarH = hRow * 0.7;
                const barOffsetY = (hRow - currentBarH) / 2;
                const rowFontSize = isFourThree ? 8 : 8.5;

                if (isMilestoneTask(row)) {
                    // Milestone star
                    pptText(slide, '*', bx - 0.04, y + barOffsetY - 0.02, 0.08, currentBarH + 0.04, {
                        fontSize: isFourThree ? 11 : 13,
                        bold: true,
                        color: style.barColor1,
                        align: 'center',
                        margin: 0,
                    });
                    // Milestone label to the right of the star
                    const labelX = bx + 0.06;
                    const labelW = slideW - leftM - labelX;
                    pptText(slide, displaySummary, labelX, y + barOffsetY, labelW, currentBarH, {
                        fontSize: PPT_FONT_RULES.milestoneLabel,
                        bold: true,
                        color: PPT_COLORS.milestoneText,
                        align: 'left',
                    });
                } else if (isStepMeetingTask(row)) {
                    // Step meeting diamond shape
                    const diamondSize = Math.max(0.06, hRow - 0.10);
                    const diamondOffsetY = (hRow - diamondSize) / 2;
                    pptRect(slide, shapes.diamond, bx - diamondSize / 2, y + diamondOffsetY, diamondSize, diamondSize, style.barColor1);
                    pptStepMeetingLabelBefore(slide, displaySummary, bx, y + barOffsetY, currentBarH, PPT_FONT_RULES.stepMeetingLabel, leftM);
                } else {
                    // Task bar
                    pptRect(slide, shapes.rect, bx, y + barOffsetY, bw, currentBarH, style.barColor2);
                    
                    // Smart label placement (inside vs outside)
                    const estCharW = 0.065;
                    const textW = displaySummary.length * estCharW;
                    let fits = false;
                    if (!isInfPpap) {
                        if (bw >= textW + 0.15) {
                            fits = true;
                        }
                    }

                    if (fits) {
                        // Fit inside
                        pptText(slide, displaySummary, bx, y + barOffsetY, bw, currentBarH, {
                            fontSize: rowFontSize,
                            color: 'FFFFFF',
                            bold: true,
                            align: 'center',
                        });
                    } else {
                        // Fit outside to the right with a small space or arrow
                        const labelText = (isInfPpap ? '<- ' : '') + displaySummary;
                        const labelX = bx + bw + 0.05;
                        const labelW = slideW - leftM - labelX;
                        pptText(slide, labelText, labelX, y + barOffsetY, labelW, currentBarH, {
                            fontSize: rowFontSize,
                            color: '1F2937',
                            align: 'left',
                        });
                    }
                }
            }
        });

        // 4. Draw Grouping blocks over the row backgrounds
        groupBlocks.forEach(block => {
            const by = topM + block.startIndex * hRow;
            const bh = block.count * hRow;
            const bx = leftM;
            const bw = wKey;
            
            const style = getGroupStyle(block.group);

            // Draw rounded rectangle for the grouping block with vertical gap and no border
            const blockGapY = Math.min(0.02, hRow * 0.1);
            const blockFontSize = 14;
            const displayGroupName = getDisplayGroupingName(block.group);
            pptRect(slide, shapes.roundRect, bx, by + blockGapY, bw, bh - blockGapY * 2, style.groupBlock, null);

            pptText(slide, displayGroupName, bx + 0.03, by + blockGapY, bw - 0.06, bh - blockGapY * 2, {
                fontSize: blockFontSize,
                bold: true,
                color: style.groupText,
                align: 'center',
            });
        });

        drawPptTodayLine(slide, shapes, bs, be, chartL, chartW, totalDays, topM - hMonth, page.length * hRow + hMonth);

        
    });

    await pptx.writeFile({ fileName: filename });
}

// ========================================
// Export Functions
// ========================================
function getPptExportNames(tasks, layout = 'wide') {
    const stamp = new Date()
        .toISOString()
        .replace(/[-:T]/g, '')
        .slice(0, 12);
    const keyName = [...new Set(tasks
        .map(t => String(t.key || '').split('-')[0])
        .filter(Boolean))]
        .slice(0, 2)
        .join('_') || 'KEY';
    const suffix = layout === '4:3' ? '_4x3' : '';
    const base = `GanttChart_${stamp}_${keyName}${suffix}`;
    return { title: base, filename: `${base}.pptx` };
}

async function exportPPTX(layout = 'wide') {
    const tasks = state.filteredTasks.filter(t => !t.isGroup);
    if (tasks.length === 0) {
        alert('No visible tasks to export.');
        return;
    }
    
    const isFourThree = layout === '4:3';
    const { title, filename } = getPptExportNames(tasks, layout);
    
    // Prepare task data for server
    const taskData = tasks.map(t => ({
        key: t.key,
        displayKey: t.displayKey || t.grouping || t.key,
        grouping: t.grouping || '',
        type: t.type,
        summary: t.summary,
        status: t.status,
        assignee: t.assignee || '',
        parent: t.parent || '',
        startDate: t.startDate ? formatDate(t.startDate) : '',
        endDate: t.endDate ? formatDate(t.endDate) : '',
    }));
    
    // Show loading state on button
    const btn = document.getElementById(isFourThree ? 'btn-export-pptx-4x3' : 'btn-export-pptx');
    const origHTML = btn.innerHTML;
    btn.innerHTML = '<span class="btn-spinner"></span> Generating...';
    btn.disabled = true;
    
    try {
        // Include custom timeline bounds if set
        const payload = { tasks: taskData, title, filename, groupBy: state.groupBy, layout, hideSummary: state.hideSummaryColumn };
        if (state.customStart) payload.timelineStart = formatDate(state.customStart);
        if (state.customEnd) payload.timelineEnd = formatDate(state.customEnd);

        if (!isFourThree && isLocalServer()) {
            try {
                const response = await fetch('/api/generate-pptx', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload),
                });

                if (response.ok) {
                    const result = await response.json();

                    if (result.success) {
                        showNotification(`PPTX saved to output folder:\n${result.filename}`, 'success');
                        return;
                    }

                    console.warn('Server PPTX export failed:', result.error);
                }
            } catch (serverErr) {
                console.warn('Server PPTX export skipped:', serverErr);
            }
        }

        await generatePPTXInBrowser(tasks, title, filename, { layout, hideSummary: state.hideSummaryColumn });
        showNotification(`PPTX downloaded:\n${filename}`, 'success');
    } catch (err) {
        console.error('PPTX export failed:', err);
        showNotification(`PPTX export failed: ${err.message}`, 'error');
    } finally {
        btn.innerHTML = origHTML;
        btn.disabled = false;
    }
}

// ========================================
// Notification
// ========================================
function showNotification(message, type = 'success') {
    // Remove existing notification
    const existing = document.getElementById('notification');
    if (existing) existing.remove();
    
    const notif = document.createElement('div');
    notif.id = 'notification';
    notif.className = `notification notification-${type}`;
    notif.innerHTML = `
        <div class="notif-icon">
            ${type === 'success' 
                ? '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>'
                : '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>'
            }
        </div>
        <div class="notif-text">${message}</div>
        <button class="notif-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    document.body.appendChild(notif);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (notif.parentElement) {
            notif.classList.add('notification-hide');
            setTimeout(() => notif.remove(), 300);
        }
    }, 5000);
}

// ========================================
// Event Handlers
// ========================================
function initEventListeners() {
    // File input
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    
    dropZone.addEventListener('drop', async (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.csv'));
        await handleFiles(files);
    });
    
    fileInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        await handleFiles(files);
    });
    
    // View mode buttons
    document.querySelectorAll('.view-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.viewMode = btn.dataset.view;
            renderGantt();
        });
    });
    
    // Group by
    document.getElementById('group-by').addEventListener('change', (e) => {
        state.groupBy = e.target.value;
        renderGantt();
    });
    
    // Hide summary checkbox
    const hideSummaryCheckbox = document.getElementById('hide-summary');
    if (hideSummaryCheckbox) {
        state.hideSummaryColumn = hideSummaryCheckbox.checked;
        hideSummaryCheckbox.addEventListener('change', (e) => {
            state.hideSummaryColumn = e.target.checked;
            renderGantt();
        });
    }
    
    // Project select
    document.getElementById('project-select').addEventListener('change', (e) => {
        state.selectedProject = e.target.value;
        renderGantt();
    });
    
    // Export buttons
    document.getElementById('btn-export-pptx').addEventListener('click', () => exportPPTX('wide'));
    document.getElementById('btn-export-pptx-4x3').addEventListener('click', () => exportPPTX('4:3'));
    document.getElementById('btn-download-simple-csv').addEventListener('click', downloadSimpleCSV);
    
    // Timeline date range controls
    document.getElementById('timeline-start').addEventListener('change', (e) => {
        state.customStart = e.target.value ? new Date(e.target.value) : null;
        renderGantt();
    });
    
    document.getElementById('timeline-end').addEventListener('change', (e) => {
        state.customEnd = e.target.value ? new Date(e.target.value) : null;
        renderGantt();
    });
    
    document.getElementById('btn-reset-range').addEventListener('click', () => {
        state.customStart = null;
        state.customEnd = null;
        document.getElementById('timeline-start').value = '';
        document.getElementById('timeline-end').value = '';
        renderGantt();
    });
    
    // Theme toggle
    document.getElementById('btn-toggle-theme').addEventListener('click', () => {
        state.theme = state.theme === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', state.theme);
    });
}

async function handleFiles(files) {
    const loadedFilesEl = document.getElementById('loaded-files');
    
    for (const file of files) {
        try {
            const tasks = await loadCSVFile(file);
            state.tasks.push(...tasks);
            state.files.push(file.name);
            
            // Show loaded file indicator
            const fileEl = document.createElement('div');
            fileEl.className = 'loaded-file';
            fileEl.innerHTML = `
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>
                <span>${file.name} (${tasks.length} tasks)</span>
            `;
            loadedFilesEl.appendChild(fileEl);
        } catch (err) {
            console.error(`Failed to load ${file.name}:`, err);
        }
    }
    
    if (state.tasks.length > 0) {
        updateSimpleCsvDownloadButton();

        // Populate project selector
        const projects = [...new Set(state.tasks.map(t => t.project))];
        const select = document.getElementById('project-select');
        select.innerHTML = '<option value="all">All</option>';
        for (const proj of projects) {
            select.innerHTML += `<option value="${proj}">${proj}</option>`;
        }
        
        // Render after short delay for animation
        setTimeout(() => {
            document.getElementById('upload-section').style.minHeight = 'auto';
            renderGantt();
        }, 500);
    }
}

// ========================================
// Initialize
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    // Set default light theme
    document.documentElement.setAttribute('data-theme', 'light');
    initEventListeners();
    createTooltip();
});


