import React, { useState, useCallback, useMemo } from 'react';
import { Calendar, User, FileText, Copy, Download, Upload, RefreshCw, Info, Settings, Mail, Link2, ClipboardList } from 'lucide-react';

// Attendance Policy Rules Engine
const AttendancePolicyEngine = {
  sickKeywords: [
    'sick', 'ill', 'fever', 'cough', 'cold', 'flu', 'nausea', 'vomit', 
    'headache', 'migraine', 'stomach', 'broken', 'injury', 'injured',
    'doctor', 'hospital', 'clinic', 'health', 'medical', 'unwell',
    'not feeling well', "don't feel well", 'feeling unwell', 'throat',
    'sore', 'ache', 'pain', 'dizzy', 'diarrhea', 'covid', 'virus'
  ],
  
  academicKeywords: ['prelim', 'exam', 'final', 'finals', 'midterm', 'test'],
  
  isWeekendMorningShift(date, shiftTime) {
    const dayOfWeek = date.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    if (!isWeekend) return false;
    const timeMatch = shiftTime.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (!timeMatch) return false;
    let hour = parseInt(timeMatch[1]);
    const ampm = timeMatch[3]?.toLowerCase();
    if (ampm === 'pm' && hour !== 12) hour += 12;
    if (ampm === 'am' && hour === 12) hour = 0;
    return hour < 14;
  },
  
  isTuesdayOrThursday(date) {
    const day = date.getDay();
    return day === 2 || day === 4;
  },
  
  isEveningShift(shiftTime) {
    const timeMatch = shiftTime.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (!timeMatch) return false;
    let hour = parseInt(timeMatch[1]);
    const ampm = timeMatch[3]?.toLowerCase();
    if (ampm === 'pm' && hour !== 12) hour += 12;
    if (ampm === 'am' && hour === 12) hour = 0;
    return hour >= 16;
  },
  
  calculateHoursNotice(requestedDate, shiftDate, shiftTime) {
    const timeMatch = shiftTime.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);
    if (!timeMatch) return 999;
    let hour = parseInt(timeMatch[1]);
    const minutes = parseInt(timeMatch[2] || '0');
    const ampm = timeMatch[3]?.toLowerCase();
    if (ampm === 'pm' && hour !== 12) hour += 12;
    if (ampm === 'am' && hour === 12) hour = 0;
    const shiftDateTime = new Date(shiftDate);
    shiftDateTime.setHours(hour, minutes, 0, 0);
    const diffMs = shiftDateTime - requestedDate;
    return diffMs / (1000 * 60 * 60);
  },
  
  determineInfractionType(entry, options = {}) {
    const { shiftDate, shiftTime, requestedDate, comment, isPickup, isHostShift } = entry;
    const nsCWindow = typeof options.nsCWindow === 'number' ? options.nsCWindow : 48;
    const allowAnyDay2Days = !!options.allowAnyDay2Days;

    if (isPickup) {
      return {
        type: isHostShift ? 'WO Host' : 'WO',
        reason: isHostShift ? 'Door shift covered' : 'Shift covered by another employee',
        points: -1
      };
    }
    
    const commentLower = (comment || '').toLowerCase();
    const hoursNotice = this.calculateHoursNotice(requestedDate, shiftDate, shiftTime);
    
    const hasPrelim = this.academicKeywords.some(kw => commentLower.includes(kw));
    const hasFinal = commentLower.includes('final');
    
    if (hasPrelim || hasFinal) {
      if (hasFinal) {
        return { type: 'Prelim', reason: 'Final exam - excused absence', points: 0 };
      }
      if (this.isTuesdayOrThursday(shiftDate) && this.isEveningShift(shiftTime)) {
        return { type: 'Prelim', reason: 'Prelim exam - excused absence (Tue/Thu evening)', points: 0 };
      }
    }
    
    const isSick = this.sickKeywords.some(kw => commentLower.includes(kw));
    
    if (isSick) {
      const isWeekendMorning = this.isWeekendMorningShift(shiftDate, shiftTime);
      if (hoursNotice >= 2 || isWeekendMorning) {
        return { type: 'NS/S', reason: `Sick callout with ${hoursNotice.toFixed(1)} hours notice`, points: 0 };
      } else {
        return { type: 'NS/LS', reason: `Late sick callout - only ${hoursNotice.toFixed(1)} hours notice`, points: 1 };
      }
    }
    
    if (!requestedDate || hoursNotice < 0) {
      return { type: 'NS/NC', reason: 'No show / No call - no prior notice given', points: 3 };
    }

    if (allowAnyDay2Days) {
      const reqDay = new Date(requestedDate.getFullYear(), requestedDate.getMonth(), requestedDate.getDate());
      const shiftDay = new Date(shiftDate.getFullYear(), shiftDate.getMonth(), shiftDate.getDate());
      const diffDays = (shiftDay - reqDay) / (1000 * 60 * 60 * 24);
      if (diffDays >= 2) {
        return { type: 'NS/C', reason: `Called out ${diffDays.toFixed(0)} days before shift (2-day rule)`, points: 1 };
      }
    }

    if (hoursNotice >= nsCWindow) {
      return { type: 'NS/C', reason: `Called out with ${hoursNotice.toFixed(1)} hours notice (>${nsCWindow}h)`, points: 1 };
    } else {
      return { type: 'NS/LC', reason: `Late callout - only ${hoursNotice.toFixed(1)} hours notice (<${nsCWindow}h)`, points: 2 };
    }
  }
};

// Parser for W2W pages
const W2WParser = {
  parsePickupPage(text) {
    const entries = [];
    const fullText = text.replace(/\n/g, ' ').replace(/\s+/g, ' ');
    const pickupPattern = /([A-Z][a-z]+\s+[A-Z][a-z'-]+(?:\s+[A-Z][a-z'-]+)?)\s+((?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,?\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{4})\s+(\d{1,2}:?\d{0,2}\s*[ap]?m?\s*-\s*\d{1,2}:?\d{0,2}\s*[ap]?m?)\s*([\w\s-]*)/gi;
    
    let match;
    const skipWords = ['unassigned', 'approve', 'reject', 'comment', 'pickup', 'request', 'from', 'through'];
    
    while ((match = pickupPattern.exec(fullText)) !== null) {
      const name = match[1].trim();
      const dateStr = match[2].trim();
      const timeStr = match[3].trim();
      const position = match[4]?.trim().toLowerCase() || '';

      // Check if name contains skip words - if so, retry from after the first word
      const nameParts = name.toLowerCase().split(/\s+/);
      if (nameParts.some(part => skipWords.includes(part) || skipWords.some(w => part.includes(w)))) {
        // Reset to after the first word so we can try matching the real name
        pickupPattern.lastIndex = match.index + name.split(/\s+/)[0].length + 1;
        continue;
      }
      
      const parsedDate = new Date(dateStr);
      if (isNaN(parsedDate)) continue;
      
      const isHostShift = position.includes('host') || position.includes('door');
      
      entries.push({
        name: this.formatName(name),
        shiftDate: parsedDate,
        shiftTime: timeStr,
        requestedDate: new Date(),
        comment: 'Shift pickup',
        isPickup: true,
        isHostShift: isHostShift,
        rawText: match[0]
      });
    }
    
    return entries;
  },
  
  parseCalloffPage(text) {
    const entries = [];
    const fullText = text.replace(/\n/g, ' ').replace(/\s+/g, ' ');
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
      if (skipWords.some(w => m.name.toLowerCase().includes(w))) continue;
      
      const parsedDate = new Date(m.dateStr);
      if (isNaN(parsedDate)) continue;
      
      const key = `${m.name}|${parsedDate.toISOString().split('T')[0]}`;
      if (processed.has(key)) continue;
      processed.add(key);
      
      const endIdx = matches[i + 1] ? matches[i + 1].index : fullText.length;
      const entryText = fullText.substring(m.endIndex, endIdx);
      
      const timeMatch = entryText.match(/(\d{1,2}:?\d{0,2}\s*[ap]m?\s*-\s*\d{1,2}:?\d{0,2}\s*[ap]m?)/i);
      const shiftTime = timeMatch ? timeMatch[1] : '5:15pm - 9:05pm';
      
      let comment = '';
      const commentMatch = entryText.match(/(?:Published\s+)?(?:\d+\s+)?([A-Za-z][\w\s,'()/-]+?)(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d/i);
      if (commentMatch) {
        comment = commentMatch[1].trim()
          .replace(/Comment to include.*/i, '')
          .replace(/Choose if want.*/i, '')
          .trim();
      }
      
      let requestedDate = new Date();
      const reqMatch = entryText.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{1,2})\s*-\s*(\d{1,2}):(\d{2})([ap])/i);
      if (reqMatch) {
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const monthIndex = monthNames.findIndex(mn => mn.toLowerCase() === reqMatch[1].toLowerCase());
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
        isHostShift: false,
        rawText: entryText.substring(0, 200)
      });
    }
    
    return entries;
  },
  
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

// Sheet Parser for Email Generation
const SheetParser = {
  parseNSCLog(text) {
    const lines = text.trim().split('\n');
    if (lines.length < 2) return { dates: [], employees: {} };

    const headerParts = lines[0].split('\t');
    // Remove BOM from first cell if present
    if (headerParts[0]) {
      headerParts[0] = headerParts[0].replace(/^\ufeff/, '');
    }
    // Check if first column is a date (NS-C Log starts with dates, no "Name" header)
    const firstCell = headerParts[0]?.trim();
    const isFirstCellDate = firstCell && /^\d{1,2}\/\d{1,2}\/\d{4}/.test(firstCell);

    // If first cell is a date, use all headers; otherwise skip the name header column
    const dates = (isFirstCellDate ? headerParts : headerParts.slice(1))
      .map(d => d.trim())
      .filter(d => d);

    const employees = {};

    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split('\t');
      const name = parts[0]?.trim();
      if (!name) continue;

      employees[name] = { name, infractions: [], woDates: [] };

      for (let j = 1; j < parts.length && j - 1 < dates.length; j++) {
        const value = parts[j]?.trim();
        if (value && ['NS/C', 'NS/LC', 'NS/NC', 'NS/S', 'NS/LS'].includes(value)) {
          employees[name].infractions.push({ date: dates[j - 1], type: value });
        }
        // Also extract WO dates from NS-C Log for matching
        if (value && (value === 'WO' || value === 'WO Host')) {
          employees[name].woDates.push({ date: dates[j - 1], type: value });
        }
      }
    }

    return { dates, employees };
  },
  
  parseInfractionList(text) {
    const lines = text.trim().split('\n');
    if (lines.length < 2) return {};
    
    const employees = {};
    
    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split('\t');
      const name = parts[0]?.trim();
      if (!name || name === 'a') continue;
      
      employees[name] = {
        name,
        nsc: parseInt(parts[1]) || 0,
        nslc: parseInt(parts[2]) || 0,
        nsnc: parseInt(parts[3]) || 0,
        nss: parseInt(parts[4]) || 0,
        nsls: parseInt(parts[5]) || 0,
        late: parseFloat(parts[6]) || 0,
        wo: parseInt(parts[7]) || 0,
        woHost: parseInt(parts[8]) || 0,
        coupon: parseInt(parts[9]) || 0,
        prelim: parseInt(parts[10]) || 0,
        break: parseInt(parts[11]) || 0,
        total: parseInt(parts[12]) || 0
      };
    }
    
    return employees;
  },
  
  parseWOExpiration(text) {
    const lines = text.trim().split('\n');
    const fullShifts = [];
    const partialShifts = [];
    
    for (let i = 0; i < lines.length; i++) {
      const parts = lines[i].split('\t');
      
      if (parts[0]?.toLowerCase().includes('name') || 
          parts[0]?.toLowerCase().includes('full') ||
          parts[0]?.toLowerCase().includes('partial')) continue;
      
      const fullName = parts[0]?.trim();
      const fullWODate = parts[1]?.trim();
      
      if (fullName && fullWODate) {
        fullShifts.push({ name: fullName, woDate: fullWODate });
      }
      
      const partialName = parts[6]?.trim();
      const partialWODate = parts[7]?.trim();
      
      if (partialName && partialWODate) {
        partialShifts.push({ name: partialName, woDate: partialWODate });
      }
    }
    
    return { fullShifts, partialShifts };
  }
};

// Normalization Helper for matching names across different sheets
const normalizeName = (name) => {
  if (!name) return '';
  // Remove trailing tags like (HOST) or (STU)
  let clean = name.replace(/\s*\([^)]*\)\s*/g, '').toLowerCase().trim();
  // Handle "Surname, First" -> "First Surname"
  if (clean.includes(',')) {
    const parts = clean.split(',').map(s => s.trim());
    return `${parts[1]} ${parts[0]}`;
  }
  return clean;
};

function parseDate(dateStr) {
  if (!dateStr) return null;
  // Strip day-of-week suffix like "12/09/2025 Tue" -> "12/09/2025"
  const cleaned = dateStr.toString().replace(/\s+(Sun|Mon|Tue|Wed|Thu|Fri|Sat).*$/i, '').trim();
  const d = new Date(cleaned);
  return isNaN(d) ? null : d;
}

function isWithinDays(date1, date2, days) {
  const d1 = parseDate(date1);
  const d2 = parseDate(date2);
  if (!d1 || !d2) return false;
  d1.setHours(0,0,0,0);
  d2.setHours(0,0,0,0);
  // WO can be within 14 days before or after the infraction
  const diffTime = Math.abs(d1.getTime() - d2.getTime());
  const diffDays = diffTime / (1000 * 60 * 60 * 24);
  return diffDays <= days;
}

function formatDateShort(dateStr) {
  if (!dateStr) return '';
  const d = parseDate(dateStr);
  if (!d) return dateStr;
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()} ${days[d.getDay()]}`;
}

function getFirstName(fullName) {
  if (!fullName) return '';
  const parts = fullName.split(',');
  return parts.length >= 2 ? parts[1].trim().split(' ')[0] : fullName.split(' ')[0];
}

function getInfractionDisplayName(type) {
  return { 'NS/C': 'No Show/Call', 'NS/LC': 'No Show/Late Call', 'NS/NC': 'No Show/No Call', 
           'NS/S': 'Sick Call (excused)', 'NS/LS': 'Late Sick Call' }[type] || type;
}

function getInfractionPoints(type) {
  return { 'NS/C': 1, 'NS/LC': 2, 'NS/NC': 3, 'NS/S': 0, 'NS/LS': 1 }[type] || 0;
}

export default function W2WAttendanceProcessor() {
  const [pickupText, setPickupText] = useState('');
  const [calloffText, setCalloffText] = useState('');
  const [processedEntries, setProcessedEntries] = useState([]);
  const [activeTab, setActiveTab] = useState('input');
  const [showInstructions, setShowInstructions] = useState(false);
  const [copySuccess, setCopySuccess] = useState('');
  const [managerName, setManagerName] = useState('MANAGER');
  const [nsCWindow, setNsCWindow] = useState(48);
  const [allowAnyDay2Days, setAllowAnyDay2Days] = useState(false);

  const [nscLogText, setNscLogText] = useState('');
  const [infractionListText, setInfractionListText] = useState('');
  const [woExpirationText, setWoExpirationText] = useState('');
  const [emailData, setEmailData] = useState([]);
  const [selectedEmployee, setSelectedEmployee] = useState(null);

  const processData = useCallback(() => {
    const allEntries = [];
    if (pickupText.trim()) allEntries.push(...W2WParser.parsePickupPage(pickupText));
    if (calloffText.trim()) allEntries.push(...W2WParser.parseCalloffPage(calloffText));
    
    const processed = allEntries.map((entry, idx) => {
      const infraction = AttendancePolicyEngine.determineInfractionType(entry, { nsCWindow, allowAnyDay2Days });
      return { ...entry, id: idx, infraction: infraction.type, reason: infraction.reason, 
               points: infraction.points, isCancelled: false };
    });
    
    processed.sort((a, b) => a.shiftDate - b.shiftDate);
    
    const infractions = processed.filter(e => ['NS/C', 'NS/LC', 'NS/NC'].includes(e.infraction));
    const workOffs = processed.filter(e => ['WO', 'WO Host'].includes(e.infraction));
    
    infractions.forEach(inf => {
      const matchingWo = workOffs.find(wo => wo.name === inf.name && !wo.isCancelled && isWithinDays(wo.shiftDate, inf.shiftDate, 14));
      if (matchingWo) {
        inf.isCancelled = true;
        matchingWo.isCancelled = true;
      }
    });
    
    setProcessedEntries(processed);
    setActiveTab('results');
  }, [pickupText, calloffText, nsCWindow, allowAnyDay2Days]);

  const processEmailSheets = useCallback(() => {
    const nscLog = SheetParser.parseNSCLog(nscLogText);
    const infractionList = SheetParser.parseInfractionList(infractionListText);
    const woExpiration = SheetParser.parseWOExpiration(woExpirationText);

    // Create normalized lookups to handle name variations
    const normalizedWos = {};
    // Add WOs from WO Expiration sheet
    [...woExpiration.fullShifts, ...woExpiration.partialShifts].forEach(wo => {
      const norm = normalizeName(wo.name);
      if (!normalizedWos[norm]) normalizedWos[norm] = [];
      normalizedWos[norm].push(wo);
    });

    const normalizedNscLog = {};
    Object.keys(nscLog.employees).forEach(n => {
      const norm = normalizeName(n);
      normalizedNscLog[norm] = nscLog.employees[n];
      // Also add WO dates from NS-C Log (may not be in WO Expiration sheet)
      if (nscLog.employees[n].woDates) {
        if (!normalizedWos[norm]) normalizedWos[norm] = [];
        nscLog.employees[n].woDates.forEach(wo => {
          // Avoid duplicates - check if this date already exists
          const woDateParsed = parseDate(wo.date);
          if (!woDateParsed) return;
          const exists = normalizedWos[norm].some(existing => {
            const existingParsed = parseDate(existing.woDate);
            if (!existingParsed) return false;
            return woDateParsed.setHours(0,0,0,0) === existingParsed.setHours(0,0,0,0);
          });
          if (!exists) {
            normalizedWos[norm].push({ name: n, woDate: wo.date });
          }
        });
      }
    });

    const employeeEmails = [];

    Object.keys(infractionList).forEach(rawName => {
      const normName = normalizeName(rawName);
      const infData = infractionList[rawName];
      const nscData = normalizedNscLog[normName];
      let availableWos = [...(normalizedWos[normName] || [])];

      if (infData.total <= 0) return;

      const infractions = nscData?.infractions || [];
      const today = new Date();

      // Count infractions from NS-C Log by type
      const logCounts = { nsc: 0, nslc: 0, nsnc: 0, nss: 0, nsls: 0 };
      infractions.forEach(inf => {
        if (inf.type === 'NS/C') logCounts.nsc++;
        else if (inf.type === 'NS/LC') logCounts.nslc++;
        else if (inf.type === 'NS/NC') logCounts.nsnc++;
        else if (inf.type === 'NS/S') logCounts.nss++;
        else if (inf.type === 'NS/LS') logCounts.nsls++;
      });

      // Calculate how many infractions are missing from NS-C Log (historical data)
      const missingCounts = {
        nsc: Math.max(0, infData.nsc - logCounts.nsc),
        nslc: Math.max(0, infData.nslc - logCounts.nslc),
        nsnc: Math.max(0, infData.nsnc - logCounts.nsnc),
        nss: Math.max(0, infData.nss - logCounts.nss),
        nsls: Math.max(0, infData.nsls - logCounts.nsls)
      };

      const allInfractions = [];

      // First, add infractions from NS-C Log (these have dates)
      infractions.forEach(inf => {
        const infDate = parseDate(inf.date);
        const twoWeeksAfter = infDate ? new Date(infDate) : null;
        if (twoWeeksAfter) twoWeeksAfter.setDate(twoWeeksAfter.getDate() + 14);

        let pts = getInfractionPoints(inf.type);
        let status = '';

        if (pts > 0) {
          const woIndex = availableWos.findIndex(wo => isWithinDays(wo.woDate, inf.date, 14));

          if (woIndex !== -1) {
            status = ' (already made up)';
            pts = Math.max(0, pts - 1);
            availableWos.splice(woIndex, 1);
          } else if (twoWeeksAfter && today > twoWeeksAfter) {
            status = ' (can no longer make up at this time)';
          }
        }

        allInfractions.push({ ...inf, points: pts, status, hasDate: true });
      });

      // Then add missing infractions from Infraction List counts (no dates, but match with remaining WOs)
      const addMissingInfractions = (count, type, basePts) => {
        for (let i = 0; i < count; i++) {
          let pts = basePts;
          let status = '';

          if (basePts > 0) {
            // Try to match with remaining WOs (these would be older, already expired make-up windows)
            if (availableWos.length > 0) {
              status = ' (already made up)';
              pts = Math.max(0, pts - 1);
              availableWos.shift(); // Use oldest WO first
            } else {
              // No WO and no date means it's historical and can no longer be made up
              status = ' (can no longer make up at this time)';
            }
          }

          allInfractions.push({ type, points: pts, status, hasDate: false, date: null });
        }
      };

      addMissingInfractions(missingCounts.nsc, 'NS/C', 1);
      addMissingInfractions(missingCounts.nslc, 'NS/LC', 2);
      addMissingInfractions(missingCounts.nsnc, 'NS/NC', 3);
      addMissingInfractions(missingCounts.nss, 'NS/S', 0);
      addMissingInfractions(missingCounts.nsls, 'NS/LS', 1);

      // Sort: dated infractions first (by date), then undated ones
      allInfractions.sort((a, b) => {
        if (a.hasDate && b.hasDate) return new Date(a.date) - new Date(b.date);
        if (a.hasDate) return -1;
        if (b.hasDate) return 1;
        return 0;
      });

      employeeEmails.push({
        name: rawName, firstName: getFirstName(rawName), totalPoints: infData.total,
        infractions: allInfractions, counts: infData
      });
    });

    employeeEmails.sort((a, b) => a.name.localeCompare(b.name));
    setEmailData(employeeEmails);
    if (employeeEmails.length > 0) setSelectedEmployee(employeeEmails[0].name);
  }, [nscLogText, infractionListText, woExpirationText]);

  const generateEmail = useCallback((employee) => {
    if (!employee) return '';

    const lines = [
      `Hi ${employee.firstName},`,
      '',
      `This is a reminder that you have accumulated ${employee.totalPoints} infraction point${employee.totalPoints !== 1 ? 's' : ''} so far.`,
      '',
      'This includes:',
      ''
    ];

    employee.infractions.forEach(inf => {
      if (inf.points === 0) return; // Skip 0-point infractions
      if (inf.hasDate) {
        lines.push(`    ${inf.points}: ${formatDateShort(inf.date)} : ${getInfractionDisplayName(inf.type)}${inf.status}`);
      } else {
        // Historical infraction without date from NS-C Log
        lines.push(`    ${inf.points}: ${getInfractionDisplayName(inf.type)}${inf.status}`);
      }
    });

    lines.push('', 'Please note that you have 2 weeks after the call-off date to make up your call-off shifts. Let us know if you need an extension.');
    lines.push('', 'To be in good standing is to have at most 3 infraction points. Please also let us know if we make any mistakes or if you have any questions.');
    lines.push('', 'Thank you,', managerName);
    lines.push('', 'Note: This email was generated using an alpha version of our attendance tracking software. If you notice any mistakes or discrepancies, please let us know by replying to this email.');

    return lines.join('\n');
  }, [managerName]);

  const woExpirationEntries = useMemo(() => {
    const woEntries = processedEntries.filter(e => ['WO', 'WO Host'].includes(e.infraction));
    const hostByDate = {};
    woEntries.filter(e => e.infraction === 'WO Host').forEach(e => {
      const key = `${e.name}|${e.shiftDate.toISOString().split('T')[0]}`;
      if (!hostByDate[key]) hostByDate[key] = [];
      hostByDate[key].push(e);
    });
    
    const fullShifts = [], partialDoorShifts = [], processedKeys = new Set();
    
    woEntries.forEach(e => {
      const woDate = new Date(e.shiftDate);
      const entry = { name: e.name, woDate, isCancelled: e.isCancelled };
      
      if (e.infraction === 'WO') {
        fullShifts.push(entry);
      } else {
        const key = `${e.name}|${e.shiftDate.toISOString().split('T')[0]}`;
        if (hostByDate[key]?.length >= 2) {
          if (!processedKeys.has(key)) {
            entry.displayName = `${e.name} (HOST)`;
            fullShifts.push(entry);
            processedKeys.add(key);
          }
        } else {
          partialDoorShifts.push(entry);
        }
      }
    });
    
    return { fullShifts, partialDoorShifts };
  }, [processedEntries]);

  const generateGoogleSheetsScript = () => {
    if (processedEntries.length === 0) return '';
    
    const entries = processedEntries.map(e => ({
      name: e.name, date: e.shiftDate.toISOString().split('T')[0], infraction: e.infraction, isCancelled: e.isCancelled
    }));
    
    const woExp = {
      fullShifts: woExpirationEntries.fullShifts.map(e => ({
        name: e.displayName || e.name,
        woDate: e.woDate.toISOString().split('T')[0]
      })),
      partialDoorShifts: woExpirationEntries.partialDoorShifts.map(e => ({
        name: e.name,
        woDate: e.woDate.toISOString().split('T')[0]
      }))
    };
    
    return `function addInfractions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('NS-C Log');
  const woSheet = ss.getSheetByName('WO Expiration');
  
  const entries = ${JSON.stringify(entries, null, 2)};
  const woExpiration = ${JSON.stringify(woExp, null, 2)};
  
  const headerRow = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const dateColumns = {};
  for (let col = 1; col < headerRow.length; col++) {
    if (headerRow[col] instanceof Date) {
      dateColumns[Utilities.formatDate(headerRow[col], Session.getScriptTimeZone(), 'yyyy-MM-dd')] = col + 1;
    }
  }
  
  const nameColumn = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues();
  const employeeRows = {};
  nameColumn.forEach((row, i) => { if (row[0]) employeeRows[row[0].trim()] = i + 2; });
  
  let added = 0, errors = [];
  entries.forEach(e => {
    const row = employeeRows[e.name], col = dateColumns[e.date];
    if (row && col) {
      const cell = logSheet.getRange(row, col);
      if (!cell.getValue()) {
        cell.setValue(e.infraction);
        if (e.isCancelled) cell.setBackground('#FFFF00');
        added++;
      }
    } else {
      errors.push((row ? '' : 'Employee: ' + e.name) + (col ? '' : ' Date: ' + e.date));
    }
  });
  
  if (woSheet) {
    let fullRow = 3; while (woSheet.getRange(fullRow, 1).getValue()) fullRow++;
    let partialRow = 3; while (woSheet.getRange(partialRow, 7).getValue()) partialRow++;
    
    woExpiration.fullShifts.forEach(wo => {
      const parts = wo.woDate.split('-');
      const dateStr = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];
      woSheet.getRange(fullRow, 1).setValue(wo.name);
      woSheet.getRange(fullRow, 2).setValue(dateStr);
      fullRow++;
    });

    woExpiration.partialDoorShifts.forEach(wo => {
      const parts = wo.woDate.split('-');
      const dateStr = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];
      woSheet.getRange(partialRow, 7).setValue(wo.name);
      woSheet.getRange(partialRow, 8).setValue(dateStr);
      partialRow++;
    });
  }
  
  SpreadsheetApp.getUi().alert('Added ' + added + ' infractions.\\n' + 
    (errors.length ? 'Errors: ' + errors.join(', ') : 'No errors.'));
}`;
  };

  const copyToClipboard = async (text, label) => {
    try { await navigator.clipboard.writeText(text); setCopySuccess(label); setTimeout(() => setCopySuccess(''), 2000); } 
    catch (e) { console.error(e); }
  };

  const getInfractionColor = (type, isCancelled) => {
    if (isCancelled) return 'bg-yellow-200 text-yellow-900';
    return { 'NS/C': 'bg-yellow-100 text-yellow-800', 'NS/LC': 'bg-orange-100 text-orange-800',
             'NS/NC': 'bg-red-100 text-red-800', 'NS/S': 'bg-green-100 text-green-800',
             'NS/LS': 'bg-amber-100 text-amber-800', 'WO': 'bg-blue-100 text-blue-800',
             'WO Host': 'bg-indigo-100 text-indigo-800', 'Prelim': 'bg-purple-100 text-purple-800'
           }[type] || 'bg-gray-100 text-gray-800';
  };

  const selectedEmployeeData = emailData.find(e => e.name === selectedEmployee);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4">
      <div className="max-w-6xl mx-auto">
        <div className="bg-white rounded-xl shadow-lg p-6 mb-6">
          <div className="flex items-center justify-between flex-wrap gap-4">
            <div>
              <h1 className="text-2xl font-bold text-slate-800">W2W Attendance Helper</h1>
              <p className="text-slate-500 mt-1">Process requests & generate weekly emails</p>
            </div>
            <button onClick={() => setShowInstructions(!showInstructions)}
              className="flex items-center gap-2 px-4 py-2 bg-slate-100 hover:bg-slate-200 rounded-lg">
              <Info size={18} /> {showInstructions ? 'Hide' : 'Help'}
            </button>
          </div>
          
          {showInstructions && (
            <div className="mt-4 p-4 bg-blue-50 rounded-lg border border-blue-200 text-sm grid md:grid-cols-2 gap-4">
              <div>
                <p className="font-semibold text-blue-800 mb-1">Daily: Process W2W</p>
                <p className="text-blue-700">Paste pickups + calloffs → Process → Export script</p>
              </div>
              <div>
                <p className="font-semibold text-blue-800 mb-1">Weekly: Send Emails</p>
                <p className="text-blue-700">Paste 3 sheets → Generate → Click name → Copy email</p>
              </div>
            </div>
          )}
        </div>

        <div className="bg-white rounded-xl shadow-lg p-4 mb-6 flex flex-wrap gap-4 items-center">
          <label className="flex items-center gap-2 text-sm">
            NS-C: <input type="range" min={12} max={72} value={nsCWindow} onChange={e => setNsCWindow(+e.target.value)} className="w-20" />
            <span className="text-blue-600 font-semibold w-8">{nsCWindow}h</span>
          </label>
          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" checked={allowAnyDay2Days} onChange={e => setAllowAnyDay2Days(e.target.checked)} />
            Allow all call-offs to be made up within 2 days
          </label>
          <label className="flex items-center gap-2 text-sm">
            Manager: <input type="text" value={managerName} onChange={e => setManagerName(e.target.value)} className="px-2 py-1 border rounded w-24" />
          </label>
        </div>

        <div className="flex gap-2 mb-4 flex-wrap">
          {['input', 'results', 'export', 'emails'].map(tab => (
            <button key={tab} onClick={() => setActiveTab(tab)}
              className={`px-4 py-2 rounded-lg font-medium flex items-center gap-2 ${activeTab === tab ? 'bg-blue-600 text-white' : 'bg-white text-slate-600 hover:bg-slate-50'}`}>
              {tab === 'emails' ? <Mail size={16}/> : tab === 'input' ? <Upload size={16}/> : tab === 'results' ? <FileText size={16}/> : <Settings size={16}/>}
              {tab.charAt(0).toUpperCase() + tab.slice(1)}
            </button>
          ))}
        </div>

        {activeTab === 'input' && (
          <div className="space-y-4">
            <div className="grid md:grid-cols-2 gap-4">
              <div className="bg-white rounded-xl shadow-lg p-6">
                <h2 className="font-semibold text-slate-800 mb-3 flex items-center gap-2"><Upload size={18} className="text-blue-600"/> Pickups</h2>
                <textarea value={pickupText} onChange={e => setPickupText(e.target.value)} placeholder="Paste Trades Awaiting Approval..." className="w-full h-56 p-3 border rounded-lg font-mono text-sm"/>
              </div>
              <div className="bg-white rounded-xl shadow-lg p-6">
                <h2 className="font-semibold text-slate-800 mb-3 flex items-center gap-2"><Calendar size={18} className="text-orange-600"/> Calloffs</h2>
                <textarea value={calloffText} onChange={e => setCalloffText(e.target.value)} placeholder="Paste Pending Time Off..." className="w-full h-56 p-3 border rounded-lg font-mono text-sm"/>
              </div>
            </div>
            <div className="flex justify-center">
              <button onClick={processData} disabled={!pickupText.trim() && !calloffText.trim()} className="flex items-center gap-2 px-8 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-300 text-white font-semibold rounded-xl shadow-lg">
                <RefreshCw size={20}/> Let's Go
              </button>
            </div>
          </div>
        )}

        {activeTab === 'results' && (
          <div className="bg-white rounded-xl shadow-lg p-6">
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">Results ({processedEntries.length})</h2>
              <div className="flex gap-3 text-sm">
                <span>Net: {processedEntries.filter(e => !e.isCancelled).reduce((s,e) => s + e.points, 0)} pts</span>
                <span className="bg-yellow-200 px-2 py-0.5 rounded text-xs">Yellow = Matched</span>
              </div>
            </div>
            {processedEntries.length === 0 ? <p className="text-slate-400 text-center py-12">Process data first</p> : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead><tr className="border-b"><th className="text-left py-2 px-3">Employee</th><th className="text-left py-2 px-3">Date</th><th className="text-left py-2 px-3">Type</th><th className="text-left py-2 px-3">Pts</th></tr></thead>
                  <tbody>
                    {processedEntries.map((e,i) => (
                      <tr key={i} className={`border-b ${e.isCancelled ? 'bg-yellow-50' : 'hover:bg-slate-50'}`}>
                        <td className="py-2 px-3 font-medium">{e.name}</td>
                        <td className="py-2 px-3">{e.shiftDate.toLocaleDateString('en-US',{month:'short',day:'numeric'})}</td>
                        <td className="py-2 px-3"><span className={`px-2 py-0.5 rounded-full text-xs font-semibold ${getInfractionColor(e.infraction, e.isCancelled)}`}>{e.infraction}</span></td>
                        <td className="py-2 px-3"><span className={e.isCancelled ? 'line-through text-slate-400' : e.points > 0 ? 'text-red-600' : e.points < 0 ? 'text-green-600' : ''}>{e.points > 0 ? '+' : ''}{e.points}</span></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {activeTab === 'export' && (
          <div className="bg-white rounded-xl shadow-lg p-6">
            <div className="flex justify-between items-center mb-4">
              <h2 className="font-semibold">Google Sheets Script</h2>
              <button onClick={() => copyToClipboard(generateGoogleSheetsScript(), 'script')} disabled={processedEntries.length === 0} className="flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 disabled:bg-slate-300 text-white rounded-lg">
                <Copy size={16}/> {copySuccess === 'script' ? 'Copied!' : 'Copy'}
              </button>
            </div>
            <pre className="bg-slate-900 text-green-400 p-4 rounded-lg text-xs overflow-auto max-h-72 font-mono">{processedEntries.length > 0 ? generateGoogleSheetsScript() : '// Process data first'}</pre>
          </div>
        )}

        {activeTab === 'emails' && (
          <div className="space-y-4">
            <div className="bg-white rounded-xl shadow-lg p-6">
              <h2 className="font-semibold mb-4 flex items-center gap-2"><ClipboardList size={18} className="text-purple-600"/> Paste Sheets (Ctrl+A → Ctrl+C)</h2>
              <div className="grid md:grid-cols-3 gap-4 mb-4">
                <div><label className="block text-sm font-medium mb-1">NS-C Log</label><textarea value={nscLogText} onChange={e => setNscLogText(e.target.value)} placeholder="Paste NS-C Log..." className="w-full h-28 p-2 border rounded-lg font-mono text-xs"/></div>
                <div><label className="block text-sm font-medium mb-1">Infraction List</label><textarea value={infractionListText} onChange={e => setInfractionListText(e.target.value)} placeholder="Paste Infraction List..." className="w-full h-28 p-2 border rounded-lg font-mono text-xs"/></div>
                <div><label className="block text-sm font-medium mb-1">WO Expiration</label><textarea value={woExpirationText} onChange={e => setWoExpirationText(e.target.value)} placeholder="Paste WO Expiration..." className="w-full h-28 p-2 border rounded-lg font-mono text-xs"/></div>
              </div>
              <button onClick={processEmailSheets} disabled={!infractionListText.trim()} className="flex items-center gap-2 px-6 py-2 bg-purple-600 hover:bg-purple-700 disabled:bg-slate-300 text-white rounded-lg"><Mail size={18}/> Generate Emails</button>
            </div>

            {emailData.length > 0 && (
              <div className="bg-white rounded-xl shadow-lg overflow-hidden">
                <div className="flex h-96">
                  <div className="w-1/3 border-r flex flex-col min-w-0">
                    <div className="bg-slate-100 px-4 py-2 font-semibold text-sm border-b shrink-0">Employees ({emailData.length})</div>
                    <div className="overflow-y-auto flex-1">
                      {emailData.map(emp => (
                        <button key={emp.name} onClick={() => setSelectedEmployee(emp.name)} className={`w-full text-left px-4 py-3 border-b flex justify-between items-center hover:bg-slate-50 ${selectedEmployee === emp.name ? 'bg-blue-50 border-l-4 border-l-blue-600' : ''}`}>
                          <span className="font-medium">{emp.name}</span>
                          <span className={`text-xs px-2 py-0.5 rounded-full ${emp.totalPoints >= 4 ? 'bg-red-100 text-red-700' : emp.totalPoints >= 3 ? 'bg-orange-100 text-orange-700' : 'bg-yellow-100 text-yellow-700'}`}>{emp.totalPoints} pts</span>
                        </button>
                      ))}
                    </div>
                  </div>
                  <div className="w-2/3 flex flex-col min-w-0">
                    <div className="bg-slate-100 px-4 py-2 font-semibold text-sm border-b flex justify-between items-center shrink-0">
                      <span>Email Preview</span>
                      {selectedEmployeeData && <button onClick={() => copyToClipboard(generateEmail(selectedEmployeeData), 'email')} className="flex items-center gap-1 px-3 py-1 bg-purple-600 hover:bg-purple-700 text-white text-xs rounded"><Copy size={12}/> {copySuccess === 'email' ? 'Copied!' : 'Copy'}</button>}
                    </div>
                    <div className="flex-1 overflow-y-auto p-4">
                      {selectedEmployeeData ? <pre className="whitespace-pre-wrap font-sans text-sm leading-relaxed">{generateEmail(selectedEmployeeData)}</pre> : <p className="text-slate-400 text-center py-8">Select an employee</p>}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        <div className="mt-6 text-center text-xs text-slate-400">Cornell Dining • NS/C = 1pt • NS/LC = 2pts • NS/NC = 3pts</div>
      </div>
    </div>
  );
}