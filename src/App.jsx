import React, { useState, useCallback } from 'react';
import { AlertCircle, CheckCircle, Clock, Calendar, User, FileText, Copy, Download, Upload, RefreshCw, Info, Settings } from 'lucide-react';

// Attendance Policy Rules Engine
const AttendancePolicyEngine = {
  // Keywords that indicate sickness
  sickKeywords: [
    'sick', 'ill', 'fever', 'cough', 'cold', 'flu', 'nausea', 'vomit', 
    'headache', 'migraine', 'stomach', 'broken', 'injury', 'injured',
    'doctor', 'hospital', 'clinic', 'health', 'medical', 'unwell',
    'not feeling well', "don't feel well", 'feeling unwell', 'throat',
    'sore', 'ache', 'pain', 'dizzy', 'diarrhea', 'covid', 'virus'
  ],
  
  // Prelim/Final keywords
  academicKeywords: ['prelim', 'exam', 'final', 'finals', 'midterm', 'test'],
  
  // Check if a date is a weekend morning shift
  isWeekendMorningShift(date, shiftTime) {
    const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    
    if (!isWeekend) return false;
    
    // Parse shift start time
    const timeMatch = shiftTime.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (!timeMatch) return false;
    
    let hour = parseInt(timeMatch[1]);
    const ampm = timeMatch[3]?.toLowerCase();
    
    if (ampm === 'pm' && hour !== 12) hour += 12;
    if (ampm === 'am' && hour === 12) hour = 0;
    
    // Morning shift = before 2pm
    return hour < 14;
  },
  
  // Check if date is Tuesday or Thursday (for prelim detection)
  isTuesdayOrThursday(date) {
    const day = date.getDay();
    return day === 2 || day === 4; // Tuesday = 2, Thursday = 4
  },
  
  // Check if shift is evening (for prelim detection)
  isEveningShift(shiftTime) {
    const timeMatch = shiftTime.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (!timeMatch) return false;
    
    let hour = parseInt(timeMatch[1]);
    const ampm = timeMatch[3]?.toLowerCase();
    
    if (ampm === 'pm' && hour !== 12) hour += 12;
    if (ampm === 'am' && hour === 12) hour = 0;
    
    // Evening shift = 4pm or later
    return hour >= 16;
  },
  
  // Calculate hours notice given
  calculateHoursNotice(requestedDate, shiftDate, shiftTime) {
    // Parse shift time
    const timeMatch = shiftTime.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (!timeMatch) return 999; // Default to lots of notice if can't parse
    
    let hour = parseInt(timeMatch[1]);
    const minutes = parseInt(timeMatch[2] || '0');
    const ampm = timeMatch[3]?.toLowerCase();
    
    if (ampm === 'pm' && hour !== 12) hour += 12;
    if (ampm === 'am' && hour === 12) hour = 0;
    
    // Create full shift datetime
    const shiftDateTime = new Date(shiftDate);
    shiftDateTime.setHours(hour, minutes, 0, 0);
    
    // Calculate difference in hours
    const diffMs = shiftDateTime - requestedDate;
    const diffHours = diffMs / (1000 * 60 * 60);
    
    return diffHours;
  },
  
  // Determine infraction type based on all factors
  determineInfractionType(entry) {
    const { 
      shiftDate, 
      shiftTime, 
      requestedDate, 
      comment,
      isPickup 
    } = entry;
    
    // If it's a pickup (someone covering), it's a WO
    if (isPickup) {
      return {
        type: 'WO',
        reason: 'Shift covered by another employee',
        points: -1
      };
    }
    
    const commentLower = (comment || '').toLowerCase();
    const hoursNotice = this.calculateHoursNotice(requestedDate, shiftDate, shiftTime);
    
    // Check for academic exemptions
    const hasPrelim = this.academicKeywords.some(kw => commentLower.includes(kw));
    const hasFinal = commentLower.includes('final');
    
    if (hasPrelim || hasFinal) {
      // Finals can be any day
      if (hasFinal) {
        return {
          type: 'Prelim',
          reason: 'Final exam - excused absence',
          points: 0
        };
      }
      // Prelims only on Tuesday/Thursday evenings
      if (this.isTuesdayOrThursday(shiftDate) && this.isEveningShift(shiftTime)) {
        return {
          type: 'Prelim',
          reason: 'Prelim exam - excused absence (Tue/Thu evening)',
          points: 0
        };
      }
    }
    
    // Check for sick callout
    const isSick = this.sickKeywords.some(kw => commentLower.includes(kw));
    
    if (isSick) {
      const isWeekendMorning = this.isWeekendMorningShift(shiftDate, shiftTime);
      
      // NS/S requires at least 2 hours notice (except weekend mornings)
      if (hoursNotice >= 2 || isWeekendMorning) {
        return {
          type: 'NS/S',
          reason: `Sick callout with ${hoursNotice.toFixed(1)} hours notice`,
          points: 0
        };
      } else {
        // Less than 2 hours notice for sick = NS/LS
        return {
          type: 'NS/LS',
          reason: `Late sick callout - only ${hoursNotice.toFixed(1)} hours notice (requires 2+ hours)`,
          points: 1
        };
      }
    }
    
    // No show / No call - no notice given
    if (!requestedDate || hoursNotice < 0) {
      return {
        type: 'NS/NC',
        reason: 'No show / No call - no prior notice given',
        points: 3
      };
    }
    
    // Check notice period for regular callouts
    if (hoursNotice >= 48) {
      return {
        type: 'NS/C',
        reason: `Called out with ${hoursNotice.toFixed(1)} hours notice (>48 hours)`,
        points: 1
      };
    } else {
      return {
        type: 'NS/LC',
        reason: `Late callout - only ${hoursNotice.toFixed(1)} hours notice (<48 hours)`,
        points: 2
      };
    }
  }
};

// Parser for W2W pages
const W2WParser = {
  // Parse pickup request page
  parsePickupPage(text) {
    const entries = [];
    const fullText = text.replace(/\n/g, ' ').replace(/\s+/g, ' ');
    
    // Pattern: FirstName LastName Day, Month Date, Year Time - Time
    const pickupPattern = /([A-Z][a-z]+\s+[A-Z][a-z'-]+(?:\s+[A-Z][a-z'-]+)?)\s+((?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4})\s+(\d{1,2}:?\d{0,2}\s*[ap]?m?\s*-\s*\d{1,2}:?\d{0,2}\s*[ap]?m?)/gi;
    
    let match;
    const skipWords = ['unassigned', 'approve', 'reject', 'comment', 'pickup', 'request', 'from', 'through'];
    
    while ((match = pickupPattern.exec(fullText)) !== null) {
      const name = match[1].trim();
      const dateStr = match[2].trim();
      const timeStr = match[3].trim();
      
      // Skip if name contains keywords
      if (skipWords.some(w => name.toLowerCase().includes(w))) continue;
      
      const parsedDate = new Date(dateStr);
      
      if (!isNaN(parsedDate)) {
        entries.push({
          name: this.formatName(name),
          shiftDate: parsedDate,
          shiftTime: timeStr,
          requestedDate: new Date(),
          comment: 'Shift pickup',
          isPickup: true,
          rawText: match[0]
        });
      }
    }
    
    return entries;
  },
  
  // Parse time off / calloff page
  parseCalloffPage(text) {
    const entries = [];
    const fullText = text.replace(/\n/g, ' ').replace(/\s+/g, ' ');
    
    // Pattern: (Approve Deny)? FirstName LastName Day, Month Date, Year
    const calloffPattern = /(?:Approve\s+Deny\s+|Deny\s+)?([A-Z][a-z]+\s+[A-Z][a-z'-]+(?:\s+[A-Z][a-z'-]+)?)\s+((?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4})/gi;
    
    const matches = [];
    let match;
    while ((match = calloffPattern.exec(fullText)) !== null) {
      matches.push({
        name: match[1].trim(),
        dateStr: match[2].trim(),
        index: match.index,
        endIndex: match.index + match[0].length
      });
    }
    
    const skipWords = ['from', 'through', 'comment', 'requested', 'approve', 'deny', 
                       'days', 'choose', 'want', 'published', 'unassigned', 'date', 'time',
                       'student', 'supe', 'fsw', 'din', 'host', 'br'];
    const processed = new Set();
    
    for (let i = 0; i < matches.length; i++) {
      const m = matches[i];
      
      // Skip if name contains keywords
      if (skipWords.some(w => m.name.toLowerCase().includes(w))) continue;
      
      const parsedDate = new Date(m.dateStr);
      if (isNaN(parsedDate)) continue;
      
      // Create unique key
      const key = `${m.name}|${parsedDate.toISOString().split('T')[0]}`;
      if (processed.has(key)) continue;
      processed.add(key);
      
      // Get text after this match
      const endIdx = matches[i + 1] ? matches[i + 1].index : fullText.length;
      const entryText = fullText.substring(m.endIndex, endIdx);
      
      // Extract shift time
      const timeMatch = entryText.match(/(\d{1,2}:?\d{0,2}\s*[ap]m?\s*-\s*\d{1,2}:?\d{0,2}\s*[ap]m?)/i);
      const shiftTime = timeMatch ? timeMatch[1] : '5:15pm - 9:05pm';
      
      // Extract comment
      let comment = '';
      const commentMatch = entryText.match(/(?:Published\s+)?(?:\d+\s+)?([A-Za-z][\w\s,'-]+?)(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d/i);
      if (commentMatch) {
        comment = commentMatch[1].trim()
          .replace(/Comment to include.*/i, '')
          .replace(/Choose if want.*/i, '')
          .trim();
      }
      
      // Extract requested timestamp
      let requestedDate = new Date();
      const reqMatch = entryText.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{1,2})\s*-\s*(\d{1,2}):(\d{2})([ap])/i);
      if (reqMatch) {
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const monthIndex = monthNames.findIndex(m => m.toLowerCase() === reqMatch[1].toLowerCase());
        let hour = parseInt(reqMatch[3]);
        const minute = parseInt(reqMatch[4]);
        const ampm = reqMatch[5].toLowerCase();
        
        if (ampm === 'p' && hour !== 12) hour += 12;
        if (ampm === 'a' && hour === 12) hour = 0;
        
        requestedDate = new Date(2026, monthIndex, parseInt(reqMatch[2]), hour, minute);
      }
      
      entries.push({
        name: this.formatName(m.name),
        shiftDate: parsedDate,
        shiftTime: shiftTime,
        requestedDate: requestedDate,
        comment: comment,
        isPickup: false,
        rawText: entryText.substring(0, 200)
      });
    }
    
    return entries;
  },
  
  // Convert "First Last" to "Last, First" format
  formatName(name) {
    const parts = name.trim().split(/\s+/);
    if (parts.length >= 2) {
      const firstName = parts[0];
      const lastName = parts.slice(1).join(' ');
      return `${lastName}, ${firstName}`;
    }
    return name;
  }
};

// Main Component
export default function W2WAttendanceProcessor() {
  const [pickupText, setPickupText] = useState('');
  const [calloffText, setCalloffText] = useState('');
  const [processedEntries, setProcessedEntries] = useState([]);
  const [activeTab, setActiveTab] = useState('input');
  const [currentDate, setCurrentDate] = useState(new Date().toISOString().split('T')[0]);
  const [showInstructions, setShowInstructions] = useState(false);
  const [googleSheetId, setGoogleSheetId] = useState('');
  const [copySuccess, setCopySuccess] = useState('');

  const processData = useCallback(() => {
    const allEntries = [];
    
    // Parse pickup requests
    if (pickupText.trim()) {
      const pickups = W2WParser.parsePickupPage(pickupText);
      allEntries.push(...pickups);
    }
    
    // Parse calloff requests
    if (calloffText.trim()) {
      const calloffs = W2WParser.parseCalloffPage(calloffText);
      allEntries.push(...calloffs);
    }
    
    // Process each entry through the policy engine
    const processed = allEntries.map(entry => {
      const infraction = AttendancePolicyEngine.determineInfractionType(entry);
      return {
        ...entry,
        infraction: infraction.type,
        reason: infraction.reason,
        points: infraction.points
      };
    });
    
    // Sort by date
    processed.sort((a, b) => a.shiftDate - b.shiftDate);
    
    setProcessedEntries(processed);
    setActiveTab('results');
  }, [pickupText, calloffText]);

  const generateGoogleSheetsScript = () => {
    if (processedEntries.length === 0) return '';
    
    const entries = processedEntries.map(e => ({
      name: e.name,
      date: e.shiftDate.toISOString().split('T')[0],
      infraction: e.infraction
    }));
    
    return `// Google Apps Script - Run this in your Google Sheet
// Go to Extensions > Apps Script, paste this code, and run addInfractions()

function addInfractions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('NS-C Log');
  
  // Entries to add
  const entries = ${JSON.stringify(entries, null, 2)};
  
  // Get header row dates
  const headerRow = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const dateColumns = {};
  
  for (let col = 1; col < headerRow.length; col++) {
    if (headerRow[col] instanceof Date) {
      const dateStr = Utilities.formatDate(headerRow[col], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      dateColumns[dateStr] = col + 1; // +1 because arrays are 0-indexed but columns are 1-indexed
    }
  }
  
  // Get employee names
  const nameColumn = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues();
  const employeeRows = {};
  
  for (let row = 0; row < nameColumn.length; row++) {
    if (nameColumn[row][0]) {
      employeeRows[nameColumn[row][0].trim()] = row + 2; // +2 because header is row 1 and array is 0-indexed
    }
  }
  
  // Add each entry
  let added = 0;
  let errors = [];
  
  entries.forEach(entry => {
    const row = employeeRows[entry.name];
    const col = dateColumns[entry.date];
    
    if (row && col) {
      const cell = logSheet.getRange(row, col);
      const existingValue = cell.getValue();
      
      if (!existingValue) {
        cell.setValue(entry.infraction);
        added++;
      } else {
        errors.push(\`\${entry.name} on \${entry.date}: Cell already has "\${existingValue}"\`);
      }
    } else {
      if (!row) errors.push(\`Employee not found: \${entry.name}\`);
      if (!col) errors.push(\`Date not found: \${entry.date}\`);
    }
  });
  
  // Show results
  SpreadsheetApp.getUi().alert(
    \`Added \${added} infractions.\\n\\n\` +
    (errors.length > 0 ? \`Errors:\\n\${errors.join('\\n')}\` : 'No errors.')
  );
}`;
  };

  const generateCSVExport = () => {
    if (processedEntries.length === 0) return '';
    
    const headers = ['Employee Name', 'Shift Date', 'Shift Time', 'Infraction Type', 'Points', 'Reason', 'Comment'];
    const rows = processedEntries.map(e => [
      e.name,
      e.shiftDate.toLocaleDateString(),
      e.shiftTime,
      e.infraction,
      e.points,
      e.reason,
      e.comment
    ]);
    
    return [headers, ...rows].map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
  };

  const copyToClipboard = async (text, label) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopySuccess(label);
      setTimeout(() => setCopySuccess(''), 2000);
    } catch (err) {
      console.error('Failed to copy:', err);
    }
  };

  const downloadFile = (content, filename, mimeType) => {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const getInfractionColor = (type) => {
    const colors = {
      'NS/C': 'bg-yellow-100 text-yellow-800 border-yellow-300',
      'NS/LC': 'bg-orange-100 text-orange-800 border-orange-300',
      'NS/NC': 'bg-red-100 text-red-800 border-red-300',
      'NS/S': 'bg-green-100 text-green-800 border-green-300',
      'NS/LS': 'bg-amber-100 text-amber-800 border-amber-300',
      'WO': 'bg-blue-100 text-blue-800 border-blue-300',
      'Prelim': 'bg-purple-100 text-purple-800 border-purple-300'
    };
    return colors[type] || 'bg-gray-100 text-gray-800 border-gray-300';
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <div className="bg-white rounded-xl shadow-lg p-6 mb-6">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold text-slate-800">W2W Attendance Processor</h1>
              <p className="text-slate-500 mt-1">Parse WhenToWork pages and apply Cornell Dining attendance rules</p>
            </div>
            <button
              onClick={() => setShowInstructions(!showInstructions)}
              className="flex items-center gap-2 px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors"
            >
              <Info size={18} />
              {showInstructions ? 'Hide' : 'Show'} Instructions
            </button>
          </div>
          
          {showInstructions && (
            <div className="mt-4 p-4 bg-blue-50 rounded-lg border border-blue-200">
              <h3 className="font-semibold text-blue-800 mb-2">How to Use:</h3>
              <ol className="list-decimal list-inside space-y-2 text-blue-700 text-sm">
                <li>Open WhenToWork and go to <strong>Trades → Trades Awaiting Approval</strong> for pickups</li>
                <li>Select all (Ctrl+A) and copy (Ctrl+C) the entire page</li>
                <li>Paste into the <strong>Pickup Requests</strong> box below</li>
                <li>Go to <strong>Time Off → Pending</strong> for calloffs</li>
                <li>Copy and paste into the <strong>Calloff Requests</strong> box</li>
                <li>Click <strong>Process Data</strong> to analyze</li>
                <li>Review results and export to Google Sheets</li>
              </ol>
              
              <h3 className="font-semibold text-blue-800 mt-4 mb-2">Infraction Rules Applied:</h3>
              <ul className="space-y-1 text-blue-700 text-sm">
                <li><span className="font-semibold">NS/C (1pt):</span> Called out with 48+ hours notice</li>
                <li><span className="font-semibold">NS/LC (2pts):</span> Called out with less than 48 hours notice</li>
                <li><span className="font-semibold">NS/NC (3pts):</span> No show, no call</li>
                <li><span className="font-semibold">NS/S (0pts):</span> Sick with 2+ hours notice (keywords: sick, fever, cough, etc.)</li>
                <li><span className="font-semibold">NS/LS (1pt):</span> Sick with less than 2 hours notice</li>
                <li><span className="font-semibold">Prelim (0pts):</span> Tue/Thu evening shifts with "prelim" in comment, or any day for "final"</li>
                <li><span className="font-semibold">WO (-1pt):</span> Pickup/work-off shift covered</li>
              </ul>
            </div>
          )}
        </div>

        {/* Tabs */}
        <div className="flex gap-2 mb-4">
          {['input', 'results', 'export'].map(tab => (
            <button
              key={tab}
              onClick={() => setActiveTab(tab)}
              className={`px-6 py-2 rounded-lg font-medium transition-all ${
                activeTab === tab
                  ? 'bg-blue-600 text-white shadow-md'
                  : 'bg-white text-slate-600 hover:bg-slate-50'
              }`}
            >
              {tab.charAt(0).toUpperCase() + tab.slice(1)}
            </button>
          ))}
        </div>

        {/* Input Tab */}
        {activeTab === 'input' && (
          <div className="space-y-4">
            <div className="grid md:grid-cols-2 gap-4">
              {/* Pickup Requests */}
              <div className="bg-white rounded-xl shadow-lg p-6">
                <div className="flex items-center gap-2 mb-4">
                  <Upload size={20} className="text-blue-600" />
                  <h2 className="text-lg font-semibold text-slate-800">Pickup Requests</h2>
                </div>
                <textarea
                  value={pickupText}
                  onChange={(e) => setPickupText(e.target.value)}
                  placeholder="Paste the WhenToWork Trades Awaiting Approval page here..."
                  className="w-full h-64 p-3 border border-slate-200 rounded-lg font-mono text-sm resize-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
                <p className="text-xs text-slate-400 mt-2">
                  These are shift pickups (WO entries)
                </p>
              </div>

              {/* Calloff Requests */}
              <div className="bg-white rounded-xl shadow-lg p-6">
                <div className="flex items-center gap-2 mb-4">
                  <Calendar size={20} className="text-orange-600" />
                  <h2 className="text-lg font-semibold text-slate-800">Calloff Requests</h2>
                </div>
                <textarea
                  value={calloffText}
                  onChange={(e) => setCalloffText(e.target.value)}
                  placeholder="Paste the WhenToWork Pending Time Off page here..."
                  className="w-full h-64 p-3 border border-slate-200 rounded-lg font-mono text-sm resize-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
                <p className="text-xs text-slate-400 mt-2">
                  These are time-off/calloff requests (NS/C, NS/LC, NS/S, etc.)
                </p>
              </div>
            </div>

            {/* Process Button */}
            <div className="flex justify-center">
              <button
                onClick={processData}
                disabled={!pickupText.trim() && !calloffText.trim()}
                className="flex items-center gap-2 px-8 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 text-white font-semibold rounded-xl shadow-lg transition-all transform hover:scale-105 disabled:transform-none"
              >
                <RefreshCw size={20} />
                Process Data
              </button>
            </div>
          </div>
        )}

        {/* Results Tab */}
        {activeTab === 'results' && (
          <div className="bg-white rounded-xl shadow-lg p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-lg font-semibold text-slate-800">
                Processed Entries ({processedEntries.length})
              </h2>
              {processedEntries.length > 0 && (
                <div className="flex gap-2">
                  <span className="text-sm text-slate-500">
                    Total Points Impact: {processedEntries.reduce((sum, e) => sum + e.points, 0)}
                  </span>
                </div>
              )}
            </div>

            {processedEntries.length === 0 ? (
              <div className="text-center py-12 text-slate-400">
                <FileText size={48} className="mx-auto mb-4 opacity-50" />
                <p>No entries processed yet. Paste W2W data and click Process.</p>
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-slate-200">
                      <th className="text-left py-3 px-4 font-semibold text-slate-600">Employee</th>
                      <th className="text-left py-3 px-4 font-semibold text-slate-600">Shift Date</th>
                      <th className="text-left py-3 px-4 font-semibold text-slate-600">Time</th>
                      <th className="text-left py-3 px-4 font-semibold text-slate-600">Infraction</th>
                      <th className="text-left py-3 px-4 font-semibold text-slate-600">Points</th>
                      <th className="text-left py-3 px-4 font-semibold text-slate-600">Reason</th>
                    </tr>
                  </thead>
                  <tbody>
                    {processedEntries.map((entry, idx) => (
                      <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50">
                        <td className="py-3 px-4">
                          <div className="flex items-center gap-2">
                            <User size={16} className="text-slate-400" />
                            <span className="font-medium">{entry.name}</span>
                          </div>
                        </td>
                        <td className="py-3 px-4 text-slate-600">
                          {entry.shiftDate.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' })}
                        </td>
                        <td className="py-3 px-4 text-slate-600">{entry.shiftTime}</td>
                        <td className="py-3 px-4">
                          <span className={`px-3 py-1 rounded-full text-xs font-semibold border ${getInfractionColor(entry.infraction)}`}>
                            {entry.infraction}
                          </span>
                        </td>
                        <td className="py-3 px-4">
                          <span className={`font-semibold ${entry.points > 0 ? 'text-red-600' : entry.points < 0 ? 'text-green-600' : 'text-slate-400'}`}>
                            {entry.points > 0 ? '+' : ''}{entry.points}
                          </span>
                        </td>
                        <td className="py-3 px-4 text-slate-500 text-xs max-w-xs truncate" title={entry.reason}>
                          {entry.reason}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* Export Tab */}
        {activeTab === 'export' && (
          <div className="space-y-4">
            {/* Google Apps Script Export */}
            <div className="bg-white rounded-xl shadow-lg p-6">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-2">
                  <Settings size={20} className="text-green-600" />
                  <h2 className="text-lg font-semibold text-slate-800">Google Sheets Integration</h2>
                </div>
                <button
                  onClick={() => copyToClipboard(generateGoogleSheetsScript(), 'script')}
                  disabled={processedEntries.length === 0}
                  className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 disabled:bg-slate-300 text-white rounded-lg transition-colors"
                >
                  <Copy size={16} />
                  {copySuccess === 'script' ? 'Copied!' : 'Copy Script'}
                </button>
              </div>
              
              <div className="bg-slate-50 rounded-lg p-4 mb-4">
                <h3 className="font-semibold text-slate-700 mb-2">How to use:</h3>
                <ol className="list-decimal list-inside space-y-1 text-sm text-slate-600">
                  <li>Open your NS-C Log Google Sheet</li>
                  <li>Go to <strong>Extensions → Apps Script</strong></li>
                  <li>Delete any existing code and paste the script below</li>
                  <li>Click <strong>Run → addInfractions</strong></li>
                  <li>Authorize the script when prompted</li>
                  <li>The infractions will be added to the correct cells!</li>
                </ol>
              </div>

              <div className="relative">
                <pre className="bg-slate-900 text-green-400 p-4 rounded-lg text-xs overflow-auto max-h-64 font-mono">
                  {processedEntries.length > 0 ? generateGoogleSheetsScript() : '// Process some data first to generate the script'}
                </pre>
              </div>
            </div>

            {/* CSV Export */}
            <div className="bg-white rounded-xl shadow-lg p-6">
              <div className="flex items-center justify-between mb-4">
                <div className="flex items-center gap-2">
                  <Download size={20} className="text-blue-600" />
                  <h2 className="text-lg font-semibold text-slate-800">CSV Export</h2>
                </div>
                <button
                  onClick={() => downloadFile(generateCSVExport(), 'attendance_infractions.csv', 'text/csv')}
                  disabled={processedEntries.length === 0}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 text-white rounded-lg transition-colors"
                >
                  <Download size={16} />
                  Download CSV
                </button>
              </div>
              
              <p className="text-sm text-slate-500">
                Download a CSV file with all processed entries for backup or manual entry.
              </p>
            </div>

            {/* Manual Entry Reference */}
            <div className="bg-white rounded-xl shadow-lg p-6">
              <div className="flex items-center gap-2 mb-4">
                <FileText size={20} className="text-purple-600" />
                <h2 className="text-lg font-semibold text-slate-800">Quick Reference for Manual Entry</h2>
              </div>
              
              {processedEntries.length > 0 ? (
                <div className="grid gap-2">
                  {processedEntries.map((entry, idx) => (
                    <div key={idx} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                      <div className="flex items-center gap-4">
                        <span className="font-medium text-slate-700">{entry.name}</span>
                        <span className="text-slate-500">→</span>
                        <span className="text-slate-600">
                          {entry.shiftDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}
                        </span>
                      </div>
                      <span className={`px-3 py-1 rounded-full text-sm font-semibold border ${getInfractionColor(entry.infraction)}`}>
                        {entry.infraction}
                      </span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-slate-400 text-center py-4">Process data to see quick reference</p>
              )}
            </div>
          </div>
        )}

        {/* Footer */}
        <div className="mt-6 text-center text-sm text-slate-400">
          <p>Built for Cornell Dining Student Handbook compliance</p>
          <p className="mt-1">Infraction types: NS/C, NS/LC, NS/NC, NS/S, NS/LS, WO, WO Host, Prelim, Coupon</p>
        </div>
      </div>
    </div>
  );
}
