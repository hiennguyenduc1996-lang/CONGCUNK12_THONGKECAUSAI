import React, { useState, useEffect, useMemo, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI } from "@google/genai";
import { Upload, FileText, Download, Loader2, Settings, Key, Eye, EyeOff, Calculator, FlaskConical, Languages, BrainCircuit, Table as TableIcon, X, User, School, BookOpen, ChevronRight, LayoutDashboard, FileSpreadsheet, RefreshCw, ArrowUpDown, ArrowUp, ArrowDown, FileDown, Filter, Palette, Monitor, Hourglass, TrendingUp, Users, Database, Sigma, Award, Trash2, Atom, Globe, ScrollText, CheckSquare, Square, Cloud, Share2, Copy, ExternalLink, HelpCircle, Save, Link, ArrowRight, Laptop, AlertTriangle } from 'lucide-react';

// Declare libraries
declare const mammoth: any;
declare const XLSX: any;

// --- Types ---

interface StudentResult {
  sbd: string;
  name: string;
  firstName: string;
  lastName: string;
  code: string;
  rawAnswers: Record<string, string>;
  scores: {
    total: number;
    p1: number;
    p2: number;
    p3: number;
  };
  details: Record<string, 'T' | 'F'>;
}

interface QuestionStat {
  index: number;
  wrongCount: number;
  wrongPercent: number;
  correctKey: string;
}

interface GradingRow {
    stt: number;
    sbd: string;
    fullName: string;
    firstName: string; // Separate for sorting
    lastName: string;  // Separate for sorting
    class: string;
    examCode: string;
    totalScore: number;
    p1Score: number;
    p2Score: number;
    p3Score: number;
    // Map of Question Index (1-40/50) -> Student Answer (if wrong), null/empty if correct
    answers: Record<number, { val: string; isCorrect: boolean; isIgnored: boolean }>;
}

interface StudentProfile {
    id: string;
    firstName: string;
    lastName: string;
    fullName: string;
    class: string;
}

interface SubjectScores {
    math?: number;
    phys?: number;
    chem?: number;
    bio?: number;
    eng?: number;
    history?: number;
    it?: number;
    [key: string]: number | undefined;
}

// Map: ExamIndex (1-40) -> Map: StudentID -> Scores
type ExamDataStore = Record<number, Record<string, SubjectScores>>;

// --- Helper Functions ---

const formatClassName = (raw: string): string => {
    let s = String(raw || '').trim();
    // Regex: Starts with 12, followed by one or more zeros
    // 120 (1 zero) -> 12E1
    // 1200 (2 zeros) -> 12E2
    // ...
    // 120000000000 (10 zeros) -> 12E10
    const match = s.match(/^12(0+)$/);
    if (match) {
        return `12E${match[1].length}`;
    }
    return s;
};

const formatFullName = (lName: string, fName: string): string => {
    const full = `${lName} ${fName}`.replace(/\s+/g, ' ').trim();
    if (full === 'Phát Hứa Kiến') return 'Hứa Kiến Phát';
    return full;
};

const calculateGroupScore = (correctCount: number): number => {
  switch (correctCount) {
    case 1: return 0.1;
    case 2: return 0.25;
    case 3: return 0.5;
    case 4: return 1.0;
    default: return 0;
  }
};

const exportToExcel = (elementId: string, fileName: string) => {
    const table = document.getElementById(elementId);
    if (!table || typeof XLSX === 'undefined') return;
    const wb = XLSX.utils.table_to_book(table, { sheet: "ThongKe" });
    XLSX.writeFile(wb, `${fileName || 'Thong_ke'}.xlsx`);
};

// --- Generic ZipGrade Processor ---
interface GradingConfig {
    p1?: { start: number; end: number; val: number };
    p2?: { ranges: Array<[number, number]> }; // Start-End inclusive for each group
    p3?: { start: number; end: number; val: number };
    ignore?: Array<[number, number]>; // Ranges to ignore
    totalQuestions: number;
}

const processZipGradeFile = (file: File, config: GradingConfig): Promise<GradingRow[]> => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        // FORCE UTF-8 READ for CSV compatibility to fix encoding issues like "Anh LÆ°u Quá»‘c"
        reader.readAsText(file, 'UTF-8');
        
        reader.onload = (evt) => {
            try {
                const text = evt.target?.result;
                // Parse the string data
                const wb = XLSX.read(text, { type: 'string' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                
                if (data.length < 2) {
                    resolve([]);
                    return;
                }

                const headers = data[0].map((h: any) => String(h || '').trim());
                
                // Map Headers to Indices
                const mapIdx = (keys: string[]) => headers.findIndex(h => keys.some(k => h.toUpperCase() === k.toUpperCase()));
                
                const idxSBD = mapIdx(['Student ID', 'External ID', 'SBD', 'StudentID']);
                const idxFirstName = mapIdx(['First Name', 'FirstName']);
                const idxLastName = mapIdx(['Last Name', 'LastName']);
                const idxClass = mapIdx(['Class', 'Lớp']);
                const idxCode = mapIdx(['Key Version', 'Exam Code', 'Mã đề']);

                // Find Answer Columns (Stu1, PriKey1...)
                const getColIdx = (prefix: string, num: number) => headers.indexOf(`${prefix}${num}`);

                const processed: GradingRow[] = [];
                
                // Iterate rows (skip header)
                for (let r = 1; r < data.length; r++) {
                    const row = data[r];
                    if (!row || row.length === 0) continue;

                    // 1. Extract Student Info
                    const sbd = idxSBD > -1 ? String(row[idxSBD] || '') : String(row[0] || ''); 
                    
                    // Fallback for Class: User requested Column B (Index 1) if not found
                    let className = '';
                    if (idxClass > -1) className = String(row[idxClass] || '');
                    else if (row[1]) className = String(row[1] || '');
                    
                    // Format Class Name (1200 -> 12E2, etc.)
                    className = formatClassName(className);

                    // --- NAME FIX: SWAP COLUMNS ---
                    // User data often has Last Name in "First Name" col and First Name in "Last Name" col for sorting.
                    // We extract them normally first.
                    const rawFirstNameCol = idxFirstName > -1 ? String(row[idxFirstName] || '').trim() : '';
                    const rawLastNameCol = idxLastName > -1 ? String(row[idxLastName] || '').trim() : '';

                    // Then swap assignment to match Vietnamese semantics:
                    // fName (Tên) <= rawLastNameCol
                    // lName (Họ) <= rawFirstNameCol
                    const fName = rawLastNameCol; 
                    const lName = rawFirstNameCol;
                    
                    // Name Fix: Last Name + First Name + formatting
                    const fullName = formatFullName(lName, fName);

                    const code = idxCode > -1 ? String(row[idxCode] || '') : '';

                    if (!sbd && !fullName) continue;

                    // 2. Scoring
                    let p1 = 0; let p2 = 0; let p3 = 0;
                    const ansMap: Record<number, { val: string; isCorrect: boolean; isIgnored: boolean }> = {};

                    const getCellVal = (prefix: string, qNum: number) => {
                        const cIdx = getColIdx(prefix, qNum);
                        if (cIdx === -1) return '';
                        return String(row[cIdx] || '').trim().toUpperCase();
                    };

                    const isIgnored = (q: number) => {
                        if (!config.ignore) return false;
                        return config.ignore.some(([s, e]) => q >= s && q <= e);
                    };

                    // Logic P1
                    if (config.p1) {
                        for (let i = config.p1.start; i <= config.p1.end; i++) {
                            const stu = getCellVal('Stu', i);
                            const key = getCellVal('PriKey', i);
                            const correct = (stu === key && key !== '');
                            const ignored = isIgnored(i);
                            
                            if (!ignored && correct) p1 += config.p1.val;
                            ansMap[i] = { val: stu, isCorrect: correct, isIgnored: ignored };
                        }
                    }

                    // Logic P2 (Groups)
                    if (config.p2 && config.p2.ranges) {
                        config.p2.ranges.forEach(([start, end]) => {
                            let correctCount = 0;
                            const groupSize = end - start + 1;
                            for (let k = 0; k < groupSize; k++) {
                                const qIdx = start + k;
                                const stu = getCellVal('Stu', qIdx);
                                const key = getCellVal('PriKey', qIdx);
                                const correct = (stu === key && key !== '');
                                if (correct) correctCount++;
                                ansMap[qIdx] = { val: stu, isCorrect: correct, isIgnored: false };
                            }
                            p2 += calculateGroupScore(correctCount);
                        });
                    }

                    // Logic P3
                    if (config.p3) {
                        for (let i = config.p3.start; i <= config.p3.end; i++) {
                            const stu = getCellVal('Stu', i);
                            const key = getCellVal('PriKey', i);
                            const correct = (stu === key && key !== '');
                            const ignored = isIgnored(i);
                            
                            if (!ignored && correct) p3 += config.p3.val;
                            ansMap[i] = { val: stu, isCorrect: correct, isIgnored: ignored };
                        }
                    }

                    // Handle completely ignored ranges
                    for(let i=1; i <= config.totalQuestions; i++) {
                        if (!ansMap[i]) {
                            const stu = getCellVal('Stu', i);
                            ansMap[i] = { val: stu, isCorrect: false, isIgnored: true };
                        }
                    }

                    // Rounding
                    p1 = Math.round(p1 * 100) / 100;
                    p2 = Math.round(p2 * 100) / 100;
                    p3 = Math.round(p3 * 100) / 100;
                    const total = Math.round((p1 + p2 + p3) * 100) / 100;

                    processed.push({
                        stt: processed.length + 1,
                        sbd,
                        fullName,
                        firstName: fName,
                        lastName: lName,
                        class: className,
                        examCode: code,
                        totalScore: total,
                        p1Score: p1,
                        p2Score: p2,
                        p3Score: p3,
                        answers: ansMap
                    });
                }
                resolve(processed);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
    });
};

const SCRIPT_TEMPLATE = `
// Hướng dẫn:
// 1. Mở Google Sheet -> Tiện ích mở rộng -> Apps Script
// 2. Dán đoạn mã này vào file Code.gs
// 3. Nhấn "Triển khai" (Deploy) -> "Tùy chọn triển khai mới" (New deployment)
// 4. Chọn loại: "Ứng dụng web" (Web app)
// 5. Cấu hình: 
//    - Mô tả: "API Diem"
//    - Thực thi dưới dạng: "Tôi" (Me)
//    - Ai có quyền truy cập: "Bất kỳ ai" (Anyone)
// 6. Nhấn Triển khai -> Copy URL và dán vào ứng dụng.

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Đọc danh sách học sinh
  const studentSheet = ss.getSheetByName("HocSinh");
  let students = [];
  if (studentSheet) {
    const rows = studentSheet.getDataRange().getValues();
    // Bỏ qua header
    for (let i = 1; i < rows.length; i++) {
       if(rows[i][0]) {
         students.push({
           id: String(rows[i][0]),
           fullName: rows[i][1],
           firstName: rows[i][2],
           lastName: rows[i][3],
           class: rows[i][4]
         });
       }
    }
  }

  // 2. Đọc dữ liệu điểm
  const scoreSheet = ss.getSheetByName("DiemChiTiet");
  let examData = {};
  if (scoreSheet) {
     const rows = scoreSheet.getDataRange().getValues();
     for (let i = 1; i < rows.length; i++) {
        const examId = rows[i][0]; // Lần thi
        const studentId = String(rows[i][1]);
        
        if (!examData[examId]) examData[examId] = {};
        if (!examData[examId][studentId]) examData[examId][studentId] = {};
        
        // Cột: [Lần thi, SBD, Toán, Lý, Hóa, Sinh, Anh, Sử, Tin]
        // Index: 0       1    2     3   4    5     6    7    8
        if (rows[i][2] !== "") examData[examId][studentId].math = Number(rows[i][2]);
        if (rows[i][3] !== "") examData[examId][studentId].phys = Number(rows[i][3]);
        if (rows[i][4] !== "") examData[examId][studentId].chem = Number(rows[i][4]);
        if (rows[i][5] !== "") examData[examId][studentId].bio = Number(rows[i][5]);
        if (rows[i][6] !== "") examData[examId][studentId].eng = Number(rows[i][6]);
     }
  }

  const result = { students, examData };
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const body = e.postData.contents;
    const data = JSON.parse(body);

    // 1. Lưu Danh sách học sinh
    let studentSheet = ss.getSheetByName("HocSinh");
    if (!studentSheet) studentSheet = ss.insertSheet("HocSinh");
    studentSheet.clear();
    studentSheet.appendRow(["SBD", "Họ và Tên", "Tên", "Họ", "Lớp"]); // Header
    
    if (data.students && data.students.length > 0) {
      const studentRows = data.students.map(s => [
        s.id, s.fullName, s.firstName, s.lastName, s.class
      ]);
      // Write in chunks if too large, but simplified here
      studentSheet.getRange(2, 1, studentRows.length, 5).setValues(studentRows);
    }

    // 2. Lưu Điểm
    let scoreSheet = ss.getSheetByName("DiemChiTiet");
    if (!scoreSheet) scoreSheet = ss.insertSheet("DiemChiTiet");
    scoreSheet.clear();
    scoreSheet.appendRow(["Lần Thi", "SBD", "Toán", "Lý", "Hóa", "Sinh", "Anh"]);

    const scoreRows = [];
    if (data.examData) {
      Object.keys(data.examData).forEach(examTime => {
          const studentsScores = data.examData[examTime];
          Object.keys(studentsScores).forEach(sId => {
              const sc = studentsScores[sId];
              scoreRows.push([
                  examTime, 
                  sId, 
                  sc.math ?? "",
                  sc.phys ?? "",
                  sc.chem ?? "",
                  sc.bio ?? "",
                  sc.eng ?? ""
              ]);
          });
      });
    }

    if (scoreRows.length > 0) {
       scoreSheet.getRange(2, 1, scoreRows.length, 7).setValues(scoreRows);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", rowCount: scoreRows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
`;


// --- RANKING & SUMMARY COMPONENT ---

const RankingView = () => {
    const [subTab, setSubTab] = useState<'students' | 'scores' | 'summary' | 'sort-summary' | 'cloud' | 'math-grading' | 'science-grading' | 'it-grading' | 'history-grading' | 'english-grading'>('students');
    const [students, setStudents] = useState<StudentProfile[]>([]);
    const [examData, setExamData] = useState<ExamDataStore>({});
    const [activeExamTime, setActiveExamTime] = useState<number>(1);
    const [summaryTab, setSummaryTab] = useState<'math'|'phys'|'chem'|'eng'|'bio'|'A'|'A1'|'B'|'total'>('math');
    const [sortConfig, setSortConfig] = useState<{ key: string, direction: 'asc'|'desc' } | null>(null);
    
    // Cloud State
    const DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzFwQd_LBj4GrNG5P-BlmHKBO4XLrhx6aN0dutYYbTSK2GJRL04J1OSLFc09Gcc3Qlt/exec";
    const [scriptUrl, setScriptUrl] = useState(() => localStorage.getItem('gap_script_url') || DEFAULT_SCRIPT_URL);
    const [syncStatus, setSyncStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle');
    const [syncMessage, setSyncMessage] = useState('');

    // Multi-select Filter
    const [selectedClasses, setSelectedClasses] = useState<string[]>([]);
    const [isFilterOpen, setIsFilterOpen] = useState(false);
    const filterRef = useRef<HTMLDivElement>(null);

    // Grading States
    const [mathResults, setMathResults] = useState<GradingRow[]>([]);
    const [scienceResults, setScienceResults] = useState<GradingRow[]>([]);
    const [itResults, setItResults] = useState<GradingRow[]>([]);
    const [historyResults, setHistoryResults] = useState<GradingRow[]>([]);
    const [englishResults, setEnglishResults] = useState<GradingRow[]>([]);

    const [activeSortConfig, setActiveSortConfig] = useState<{ key: 'sbd' | 'name' | 'total', direction: 'asc' | 'desc' } | null>(null);

    // Click outside to close filter
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (filterRef.current && !filterRef.current.contains(event.target as Node)) {
                setIsFilterOpen(false);
            }
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

    // Save URL
    useEffect(() => {
        localStorage.setItem('gap_script_url', scriptUrl);
    }, [scriptUrl]);

    // -- Handler: Upload Student List --
    const handleStudentUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if(!file) return;

        const reader = new FileReader();
        reader.readAsBinaryString(file); 
        reader.onload = (evt) => {
            const bstr = evt.target?.result;
            const wb = XLSX.read(bstr, { type: 'binary' });
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
            const parsedStudents: StudentProfile[] = [];
            
            data.forEach((row, index) => {
                if (!row || row.length < 2) return;
                const firstCol = String(row[0] || '').trim().toLowerCase();
                if (firstCol.includes('sbd') || firstCol.includes('số báo danh')) return;

                const id = String(row[0] || '').trim();
                if (!id) return; 

                const lastName = String(row[1] || '').trim();
                const firstName = String(row[2] || '').trim();
                const cl = formatClassName(String(row[3] || ''));
                
                parsedStudents.push({
                    id,
                    firstName: firstName,
                    lastName: lastName,
                    fullName: formatFullName(lastName, firstName),
                    class: cl
                });
            });

            setStudents(parsedStudents);
            e.target.value = ''; 
        };
    };

    // -- Handler: Upload Scores (Detailed) --
    const handleScoreUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if(!file) return;

        const reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = (evt) => {
            const bstr = evt.target?.result;
            const wb = XLSX.read(bstr, { type: 'binary' });
            const sheetName = wb.SheetNames.find((n: string) => n.toUpperCase() === 'DIEMKHOI') || wb.SheetNames[0];
            const ws = wb.Sheets[sheetName];
            const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

            setExamData(prev => {
                const newData = { ...prev };
                if (!newData[activeExamTime]) newData[activeExamTime] = {};

                let count = 0;
                data.forEach((row, rowIndex) => {
                    if (rowIndex < 3) return; 
                    const id = String(row[1] || '').trim(); 
                    if (!id || id.toUpperCase() === 'SBD' || id.toUpperCase() === 'SỐ BÁO DANH') return;

                    const scores: SubjectScores = {};
                    const p = (val: any) => {
                        if (val === undefined || val === null || val === '') return undefined;
                        if (typeof val === 'number') return val;
                        const s = String(val).replace(',', '.');
                        const n = parseFloat(s);
                        return isNaN(n) ? undefined : n;
                    };

                    scores.math = p(row[5]);
                    scores.phys = p(row[6]);
                    scores.chem = p(row[7]);
                    scores.eng = p(row[8]);
                    scores.bio = p(row[9]);

                    if (Object.values(scores).some(v => v !== undefined)) {
                        newData[activeExamTime][id] = { ...(newData[activeExamTime][id] || {}), ...scores };
                        count++;
                    }
                });
                alert(`Đã tải lên điểm cho Lần ${activeExamTime}: cập nhật ${count} học sinh.`);
                return newData;
            });
            e.target.value = '';
        };
    };

    // --- GRADING HANDLERS ---
    const handleMathUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        if (!e.target.files?.[0]) return;
        const data = await processZipGradeFile(e.target.files[0], {
            totalQuestions: 40,
            p1: { start: 1, end: 12, val: 0.25 },
            ignore: [[13, 18]],
            p2: { ranges: [[19, 22], [23, 26], [27, 30], [31, 34]] },
            p3: { start: 35, end: 40, val: 0.5 }
        });
        setMathResults(data);
        e.target.value = '';
    };

    const handleScienceUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        if (!e.target.files?.[0]) return;
        const data = await processZipGradeFile(e.target.files[0], {
            totalQuestions: 40,
            p1: { start: 1, end: 18, val: 0.25 },
            p2: { ranges: [[19, 22], [23, 26], [27, 30], [31, 34]] },
            p3: { start: 35, end: 40, val: 0.25 }
        });
        setScienceResults(data);
        e.target.value = '';
    };

    const handleITUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        if (!e.target.files?.[0]) return;
        const data = await processZipGradeFile(e.target.files[0], {
            totalQuestions: 40,
            p1: { start: 1, end: 28, val: 0.25 },
            p2: { ranges: [[29, 32], [33, 36], [37, 40]] }
        });
        setItResults(data);
        e.target.value = '';
    };

    const handleHistoryUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        if (!e.target.files?.[0]) return;
        const data = await processZipGradeFile(e.target.files[0], {
            totalQuestions: 40,
            p1: { start: 1, end: 24, val: 0.25 },
            p2: { ranges: [[25, 28], [29, 32], [33, 36], [37, 40]] }
        });
        setHistoryResults(data);
        e.target.value = '';
    };

    const handleEnglishUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        if (!e.target.files?.[0]) return;
        const data = await processZipGradeFile(e.target.files[0], {
            totalQuestions: 50,
            p1: { start: 1, end: 40, val: 0.25 },
            ignore: [[41, 50]]
        });
        setEnglishResults(data);
        e.target.value = '';
    };

    const handleDeleteScore = () => {
        if(confirm(`Bạn có chắc muốn xóa dữ liệu điểm của Lần ${activeExamTime} (đưa về 0) không?`)) {
            setExamData(prev => {
                const newData = {...prev};
                if (newData[activeExamTime]) {
                    Object.keys(newData[activeExamTime]).forEach(key => {
                        newData[activeExamTime][key] = {
                            math: 0, phys: 0, chem: 0, bio: 0, eng: 0
                        };
                    });
                }
                return newData;
            });
        }
    };

    // --- CLOUD SYNC HANDLERS ---
    
    // Safety check URL to prevent crashes from passing HTML (edit page) instead of API (exec)
    const validateScriptUrl = (url: string) => {
        if (!url) return false;
        if (!url.includes('script.google.com')) return false;
        if (url.includes('/edit')) return false; // Common mistake
        if (!url.endsWith('/exec')) return false; // Should end in exec
        return true;
    };

    const handleSyncToCloud = async () => {
        if (!scriptUrl) {
            setSyncMessage("Vui lòng nhập URL Script trước.");
            setSyncStatus("error");
            return;
        }
        
        let targetUrl = scriptUrl.trim();
        if (!validateScriptUrl(targetUrl)) {
             setSyncMessage("URL không hợp lệ. URL phải kết thúc bằng '/exec' (Không phải '/edit')");
             setSyncStatus("error");
             return;
        }

        setSyncStatus("loading");
        setSyncMessage("Đang gửi dữ liệu lên Google Sheet...");
        
        try {
            const payload = {
                students: students,
                examData: examData
            };
            
            // CRITICAL FIX: Google Apps Script Web App POST requests trigger CORS preflight checks 
            // if Content-Type is 'application/json'. 
            // We MUST use 'text/plain' or 'application/x-www-form-urlencoded' to skip preflight.
            // The Headers configuration below ensures the browser treats this as a "Simple Request".
            await fetch(targetUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "text/plain;charset=utf-8",
                },
                body: JSON.stringify(payload),
            });
            
            setSyncStatus("success");
            setSyncMessage(`Đã đồng bộ thành công ${students.length} học sinh lên Sheet!`);
        } catch (error) {
            console.error(error);
            setSyncStatus("error");
            setSyncMessage("Lỗi kết nối: Không thể gửi dữ liệu. Kiểm tra URL hoặc quyền truy cập.");
        }
    };

    const handleSyncFromCloud = async () => {
        if (!scriptUrl) {
            setSyncMessage("Vui lòng nhập URL Script trước.");
            setSyncStatus("error");
            return;
        }

        let targetUrl = scriptUrl.trim();
        if (!validateScriptUrl(targetUrl)) {
             setSyncMessage("URL không hợp lệ. URL phải kết thúc bằng '/exec' (Không phải '/edit')");
             setSyncStatus("error");
             return;
        }

        setSyncStatus("loading");
        setSyncMessage("Đang tải dữ liệu từ Google Sheet...");

        try {
            const response = await fetch(targetUrl);
            if (!response.ok) {
                throw new Error("Network response was not ok");
            }
            
            const text = await response.text();
            let data: any;
            try {
                data = JSON.parse(text);
            } catch (e) {
                console.error("Failed to parse JSON", text.substring(0, 100));
                throw new Error("Dữ liệu trả về không phải JSON. Có thể URL sai hoặc quyền truy cập chưa mở.");
            }
            
            // STRICT VALIDATION TO PREVENT CRASH
            if (data && typeof data === 'object') {
                if (Array.isArray(data.students)) {
                    // Filter out null/undefined entries just in case and ensure types
                    const safeStudents = data.students
                        .filter((s: any) => s && typeof s === 'object' && s.id)
                        .map((s: any) => ({
                            ...s,
                            id: String(s.id),
                            fullName: String(s.fullName || ''),
                            firstName: String(s.firstName || ''),
                            lastName: String(s.lastName || ''),
                            class: formatClassName(String(s.class || '')) 
                        }));
                    setStudents(safeStudents);
                } else {
                     console.warn("Dữ liệu students không phải mảng", data.students);
                }

                if (data.examData && typeof data.examData === 'object') {
                    setExamData(data.examData);
                }
            } else {
                throw new Error("Cấu trúc dữ liệu không hợp lệ.");
            }
            
            const count = (data && Array.isArray(data.students)) ? data.students.length : 0;
            setSyncStatus("success");
            setSyncMessage(`Đã tải thành công: ${count} học sinh.`);
        } catch (error) {
            console.error(error);
            setSyncStatus("error");
            setSyncMessage("Lỗi: " + (error as Error).message);
        }
    };

    const getClassStats = useMemo(() => {
        const stats: Record<string, number> = {};
        if (Array.isArray(students)) {
            students.forEach(s => {
                if(s) {
                    const c = s.class || 'Khác';
                    stats[c] = (stats[c] || 0) + 1;
                }
            });
        }
        return stats;
    }, [students]);

    // Collect all unique classes from all datasets for the filter dropdown
    const uniqueClasses = useMemo(() => {
        const classes = new Set<string>();
        const safeAdd = (list: any[]) => {
            if(Array.isArray(list)) {
                list.forEach(s => s && s.class && classes.add(s.class));
            }
        };
        safeAdd(students);
        safeAdd(mathResults);
        safeAdd(scienceResults);
        safeAdd(itResults);
        safeAdd(historyResults);
        safeAdd(englishResults);
        
        return Array.from(classes).sort();
    }, [students, mathResults, scienceResults, itResults, historyResults, englishResults]);

    const getBlockType = (className: string): 'A' | 'A1' | 'B' | 'Other' => {
        const c = String(className || '').toUpperCase();
        if (c.includes('E')) return 'A1';
        if (c.includes('B')) return 'B';
        if (c.includes('A')) return 'A';
        return 'Other';
    };

    const getComputedData = useMemo(() => {
        if (!Array.isArray(students)) return [];
        const filteredStudents = selectedClasses.length > 0 ? students.filter(s => s && selectedClasses.includes(s.class)) : students;
        if (!filteredStudents.length) return [];

        const calcAvg = (values: number[]) => {
            const nonZero = values.filter(v => v !== undefined && v !== null && v !== 0);
            if (nonZero.length === 0) return 0;
            const sum = nonZero.reduce((a, b) => a + b, 0);
            return sum / nonZero.length;
        };

        const results: any[] = [];

        filteredStudents.forEach(s => {
            if (!s) return;
            const block = getBlockType(s.class);
            let shouldInclude = false;

            if (summaryTab === 'math') shouldInclude = true;
            else if (summaryTab === 'phys') shouldInclude = (block === 'A' || block === 'A1');
            else if (summaryTab === 'chem') shouldInclude = (block === 'A' || block === 'B');
            else if (summaryTab === 'bio') shouldInclude = (block === 'B');
            else if (summaryTab === 'eng') shouldInclude = (block === 'A1');
            else if (summaryTab === 'A') shouldInclude = (block === 'A');
            else if (summaryTab === 'B') shouldInclude = (block === 'B');
            else if (summaryTab === 'A1') shouldInclude = (block === 'A1');
            else if (summaryTab === 'total') shouldInclude = (block === 'A' || block === 'B' || block === 'A1');

            if (!shouldInclude) return;

            const row: any = { ...s };
            const scores = { math: [] as number[], phys: [] as number[], chem: [] as number[], bio: [] as number[], eng: [] as number[] };

            for (let i = 1; i <= 40; i++) {
                const record = examData[i]?.[s.id];
                if (record) {
                    if (record.math !== undefined) scores.math.push(record.math);
                    if (record.phys !== undefined) scores.phys.push(record.phys);
                    if (record.chem !== undefined) scores.chem.push(record.chem);
                    if (record.bio !== undefined) scores.bio.push(record.bio);
                    if (record.eng !== undefined) scores.eng.push(record.eng);
                }
            }

            const avgMath = calcAvg(scores.math);
            const avgPhys = calcAvg(scores.phys);
            const avgChem = calcAvg(scores.chem);
            const avgBio = calcAvg(scores.bio);
            const avgEng = calcAvg(scores.eng);

            let finalVal = 0;
            if (summaryTab === 'math') finalVal = avgMath;
            else if (summaryTab === 'phys') finalVal = avgPhys;
            else if (summaryTab === 'chem') finalVal = avgChem;
            else if (summaryTab === 'bio') finalVal = avgBio;
            else if (summaryTab === 'eng') finalVal = avgEng;
            else if (summaryTab === 'A') finalVal = avgMath + avgPhys + avgChem;
            else if (summaryTab === 'B') finalVal = avgMath + avgChem + avgBio;
            else if (summaryTab === 'A1') finalVal = avgMath + avgPhys + avgEng;
            else if (summaryTab === 'total') {
                if (block === 'A') finalVal = avgMath + avgPhys + avgChem;
                else if (block === 'B') finalVal = avgMath + avgChem + avgBio;
                else if (block === 'A1') finalVal = avgMath + avgPhys + avgEng;
            }

            for (let i = 1; i <= 40; i++) {
                 const record = examData[i]?.[s.id];
                 let colVal: number | undefined = undefined;
                 if (record) {
                    if (['math','phys','chem','bio','eng'].includes(summaryTab)) {
                         colVal = record[summaryTab as keyof SubjectScores];
                    } else {
                        if (summaryTab === 'A' || (summaryTab === 'total' && block === 'A')) {
                             if(record.math !== undefined && record.phys !== undefined && record.chem !== undefined) 
                                colVal = record.math + record.phys + record.chem;
                        } else if (summaryTab === 'B' || (summaryTab === 'total' && block === 'B')) {
                             if(record.math !== undefined && record.chem !== undefined && record.bio !== undefined)
                                colVal = record.math + record.chem + record.bio;
                        } else if (summaryTab === 'A1' || (summaryTab === 'total' && block === 'A1')) {
                             if(record.math !== undefined && record.phys !== undefined && record.eng !== undefined)
                                colVal = record.math + record.phys + record.eng;
                        }
                    }
                 }
                 row[`score_${i}`] = colVal;
            }

            row.avg = parseFloat(finalVal.toFixed(2));
            row.totalVal = finalVal; 
            results.push(row);
        });

        return results;
    }, [students, examData, summaryTab, selectedClasses]);

    const sortedData = useMemo(() => {
        if (!sortConfig) return getComputedData;
        const sorted = [...getComputedData];
        sorted.sort((a, b) => {
            let va = a[sortConfig.key];
            let vb = b[sortConfig.key];
            if (va === null || va === undefined) return 1;
            if (vb === null || vb === undefined) return -1;
            if (sortConfig.key === 'firstName') {
                 if (a.firstName !== b.firstName) return a.firstName.localeCompare(b.firstName) * (sortConfig.direction === 'asc' ? 1 : -1);
                 return a.lastName.localeCompare(b.lastName) * (sortConfig.direction === 'asc' ? 1 : -1);
            }
            if (va < vb) return sortConfig.direction === 'asc' ? -1 : 1;
            if (va > vb) return sortConfig.direction === 'asc' ? 1 : -1;
            return 0;
        });
        return sorted;
    }, [getComputedData, sortConfig]);

    const handleSort = (key: string) => {
        let direction: 'asc'|'desc' = 'asc';
        if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') direction = 'desc';
        setSortConfig({ key, direction });
    };

    const renderSortIcon = (key: string) => {
         if (sortConfig?.key !== key) return <ArrowUpDown size={12} style={{opacity:0.3}}/>;
         return sortConfig.direction === 'asc' ? <ArrowUp size={12}/> : <ArrowDown size={12}/>;
    };

    const handleGradingSort = (key: 'sbd' | 'name' | 'total') => {
        let direction: 'asc' | 'desc' = 'asc';
        if (activeSortConfig && activeSortConfig.key === key && activeSortConfig.direction === 'asc') {
            direction = 'desc';
        }
        setActiveSortConfig({ key, direction });
    };

    const renderGradingSortIcon = (key: 'sbd' | 'name' | 'total') => {
        if (activeSortConfig?.key !== key) return <ArrowUpDown size={12} style={{opacity:0.3}}/>;
        return activeSortConfig.direction === 'asc' ? <ArrowUp size={12}/> : <ArrowDown size={12}/>;
    };

    // Generic Sort for Grading Tables
    const getSortedGradingResults = (results: GradingRow[]) => {
        let data = [...results];
        
        // Filter by class
        if (selectedClasses.length > 0) {
            data = data.filter(r => selectedClasses.includes(r.class));
        }

        if (activeSortConfig) {
            data.sort((a, b) => {
                if (activeSortConfig.key === 'sbd') {
                    return a.sbd.localeCompare(b.sbd, 'en', { numeric: true }) * (activeSortConfig.direction === 'asc' ? 1 : -1);
                }
                if (activeSortConfig.key === 'total') {
                    return (a.totalScore - b.totalScore) * (activeSortConfig.direction === 'asc' ? 1 : -1);
                }
                if (activeSortConfig.key === 'name') {
                    const res = a.firstName.localeCompare(b.firstName, 'vi');
                    if (res !== 0) return res * (activeSortConfig.direction === 'asc' ? 1 : -1);
                    return a.lastName.localeCompare(b.lastName, 'vi') * (activeSortConfig.direction === 'asc' ? 1 : -1);
                }
                return 0;
            });
        }
        return data.map((item, index) => ({ ...item, stt: index + 1 }));
    };

    // --- FIX: Define activeExamScoreList ---
    const activeExamScoreList = useMemo(() => {
        let list = students || [];
        if (selectedClasses.length > 0) {
            list = list.filter(s => s && selectedClasses.includes(s.class));
        }
        return list.map(s => {
             if (!s) return null;
             return {
                ...s,
                scores: examData[activeExamTime]?.[s.id] || {}
            };
        }).filter(s => s !== null);
    }, [students, selectedClasses, examData, activeExamTime]);
    // ---------------------------------------

    const renderMultiSelect = () => (
        <div style={{ position: 'relative' }} ref={filterRef}>
            <button
                onClick={() => setIsFilterOpen(!isFilterOpen)}
                style={{
                    padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1', background: 'white',
                    fontSize: '13px', color: '#475569', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', minWidth: '150px'
                }}
            >
                <Filter size={14} />
                {selectedClasses.length === 0 ? "Tất cả các lớp" : `Đang chọn ${selectedClasses.length} lớp`}
            </button>
            
            {isFilterOpen && (
                <div style={{
                    position: 'absolute', top: '100%', left: 0, marginTop: '5px', background: 'white',
                    border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                    zIndex: 100, width: '250px', maxHeight: '300px', overflowY: 'auto', padding: '10px'
                }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px', fontSize: '12px' }}>
                        <span 
                            style={{ cursor: 'pointer', color: '#3b82f6', fontWeight: 600 }}
                            onClick={() => setSelectedClasses(uniqueClasses)}
                        >
                            Chọn tất cả
                        </span>
                        <span 
                            style={{ cursor: 'pointer', color: '#64748b' }}
                            onClick={() => setSelectedClasses([])}
                        >
                            Bỏ chọn
                        </span>
                    </div>
                    {uniqueClasses.map(cls => (
                        <div key={cls} style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '6px 0', fontSize: '13px' }}>
                            <input
                                type="checkbox"
                                id={`filter-${cls}`}
                                checked={selectedClasses.includes(cls)}
                                onChange={(e) => {
                                    if (e.target.checked) setSelectedClasses(prev => [...prev, cls]);
                                    else setSelectedClasses(prev => prev.filter(c => c !== cls));
                                }}
                                style={{ cursor: 'pointer' }}
                            />
                            <label htmlFor={`filter-${cls}`} style={{ cursor: 'pointer', flex: 1 }}>{cls}</label>
                        </div>
                    ))}
                    {uniqueClasses.length === 0 && <div style={{ color: '#94a3b8', fontSize: '13px' }}>Không có lớp nào</div>}
                </div>
            )}
        </div>
    );

    // Helper to render grading table
    const renderGradingTable = (results: GradingRow[], id: string, title: string, questionCount: number) => {
        const sorted = getSortedGradingResults(results);
        
        // Calculate Statistics
        const scores = sorted.map(s => s.totalScore);
        const count = scores.length;
        const max = count > 0 ? Math.max(...scores) : 0;
        const min = count > 0 ? Math.min(...scores) : 0;
        const avg = count > 0 ? (scores.reduce((a, b) => a + b, 0) / count).toFixed(2) : 0;

        return (
            <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                <div style={{ padding: '16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '20px', flexWrap: 'wrap' }}>
                     <div>
                         <h3 style={{ margin: '0 0 4px 0', color: '#1e293b' }}>{title}</h3>
                         <div style={{ display: 'flex', gap: '12px', fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                            <span>Sĩ số: <span style={{color: '#1e293b'}}>{count}</span></span>
                            <span>TB: <span style={{color: '#059669'}}>{avg}</span></span>
                            <span>Cao nhất: <span style={{color: '#059669'}}>{max}</span></span>
                            <span>Thấp nhất: <span style={{color: '#ef4444'}}>{min}</span></span>
                         </div>
                     </div>
                     
                     <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center' }}>
                        {renderMultiSelect()}
                        
                        <label style={{ 
                                padding: '10px 20px', background: '#3b82f6', color: 'white', borderRadius: '8px', fontSize: '14px', fontWeight: 600, 
                                cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                            }}>
                                <Upload size={18} /> Tải file ZipGrade
                                <input type="file" accept=".xlsx,.xls,.csv" hidden onChange={
                                    id === 'math' ? handleMathUpload :
                                    id === 'science' ? handleScienceUpload :
                                    id === 'it' ? handleITUpload :
                                    id === 'history' ? handleHistoryUpload :
                                    handleEnglishUpload
                                } />
                        </label>

                        <button
                               onClick={() => exportToExcel(`${id}-table`, `Diem_${id}`)}
                               style={{
                                  padding: '10px 20px', borderRadius: '8px', background: '#059669', border: 'none',
                                  cursor: 'pointer', fontSize: '14px', fontWeight: 600, color: 'white',
                                  display: 'flex', alignItems: 'center', gap: '8px'
                               }}
                             >
                               <FileDown size={18} /> Xuất Excel
                        </button>
                    </div>
                </div>

                <div style={{ flex: 1, overflow: 'auto' }}>
                    <table id={`${id}-table`} style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '13px', minWidth: '1500px' }}>
                        <thead style={{ position: 'sticky', top: 0, zIndex: 10, background: '#f1f5f9' }}>
                            <tr>
                                <th rowSpan={2} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', width: '40px' }}>STT</th>
                                <th rowSpan={2} onClick={() => handleGradingSort('sbd')} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', textAlign: 'left', width: '80px', cursor: 'pointer' }}>
                                    <div style={{display:'flex', alignItems:'center', gap:'4px'}}>SBD {renderGradingSortIcon('sbd')}</div>
                                </th>
                                <th rowSpan={2} onClick={() => handleGradingSort('name')} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', textAlign: 'left', minWidth: '180px', cursor: 'pointer' }}>
                                    <div style={{display:'flex', alignItems:'center', gap:'4px'}}>Họ và Tên {renderGradingSortIcon('name')}</div>
                                </th>
                                <th rowSpan={2} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', width: '60px' }}>Lớp</th>
                                <th rowSpan={2} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', width: '60px' }}>Mã đề</th>
                                <th rowSpan={2} onClick={() => handleGradingSort('total')} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', background: '#fef3c7', fontWeight: 800, color: '#b45309', cursor: 'pointer' }}>
                                    <div style={{display:'flex', alignItems:'center', gap:'4px', justifyContent: 'center'}}>Tổng {renderGradingSortIcon('total')}</div>
                                </th>
                                <th rowSpan={2} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', background: '#dbeafe' }}>P1</th>
                                <th rowSpan={2} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', background: '#dbeafe' }}>P2</th>
                                <th rowSpan={2} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', background: '#dbeafe' }}>P3</th>
                                <th colSpan={questionCount} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', textAlign: 'center' }}>Chi tiết câu hỏi</th>
                            </tr>
                            <tr>
                                {Array.from({length: questionCount}, (_, i) => i + 1).map(num => {
                                    return (
                                        <th key={num} style={{ 
                                            padding: '4px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', 
                                            width: '35px', fontSize: '11px', color: '#64748b'
                                        }}>
                                            {num}
                                        </th>
                                    );
                                })}
                            </tr>
                        </thead>
                        <tbody>
                            {sorted.length > 0 ? sorted.map((row, idx) => (
                                <tr key={idx} style={{ background: idx % 2 === 0 ? 'white' : '#fcfcfc' }}>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{row.stt}</td>
                                    <td style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', fontWeight: 600, color: '#475569' }}>{row.sbd}</td>
                                    <td style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', fontWeight: 500, whiteSpace: 'nowrap' }}>{row.fullName}</td>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{row.class}</td>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{row.examCode}</td>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', background: '#fffbeb', fontWeight: 700, color: '#b45309' }}>{row.totalScore}</td>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', background: '#eff6ff' }}>{row.p1Score}</td>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', background: '#eff6ff' }}>{row.p2Score}</td>
                                    <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', background: '#eff6ff' }}>{row.p3Score}</td>
                                    
                                    {Array.from({length: questionCount}, (_, i) => i + 1).map(num => {
                                        const ans = row.answers[num];
                                        if (ans.isIgnored) {
                                            return <td key={num} style={{ padding: '6px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', color: '#cbd5e1', background: '#f8fafc' }}>{ans.val || '-'}</td>;
                                        }
                                        if (ans.isCorrect) {
                                            return <td key={num} style={{ borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}></td>;
                                        }
                                        return (
                                            <td key={num} style={{ padding: '6px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', background: '#fee2e2', color: '#b91c1c', fontWeight: 600 }}>
                                                {ans.val}
                                            </td>
                                        );
                                    })}
                                </tr>
                            )) : (
                                <tr>
                                    <td colSpan={50} style={{ padding: '40px', textAlign: 'center', color: '#94a3b8' }}>Chưa có dữ liệu. Vui lòng tải file từ ZipGrade.</td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>
            </div>
        );
    }

    return (
        <div style={{ display: 'flex', height: '100%', background: '#f8fafc', overflow: 'hidden' }}>
            <div style={{ width: '220px', background: 'white', borderRight: '1px solid #e2e8f0', padding: '20px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <div style={{ fontSize: '12px', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', marginBottom: '10px' }}>Chức năng</div>
                <button 
                    onClick={() => setSubTab('students')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'students' ? '#eff6ff' : 'transparent',
                        color: subTab === 'students' ? '#1e3a8a' : '#64748b', fontWeight: subTab === 'students' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Users size={18} /> Danh sách học sinh
                </button>
                <button 
                    onClick={() => setSubTab('scores')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'scores' ? '#eff6ff' : 'transparent',
                        color: subTab === 'scores' ? '#1e3a8a' : '#64748b', fontWeight: subTab === 'scores' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Database size={18} /> Dữ liệu điểm
                </button>
                <button 
                    onClick={() => setSubTab('summary')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'summary' ? '#eff6ff' : 'transparent',
                        color: subTab === 'summary' ? '#1e3a8a' : '#64748b', fontWeight: subTab === 'summary' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Award size={18} /> Tổng kết
                </button>
                 <button 
                    onClick={() => setSubTab('sort-summary')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'sort-summary' ? '#eff6ff' : 'transparent',
                        color: subTab === 'sort-summary' ? '#1e3a8a' : '#64748b', fontWeight: subTab === 'sort-summary' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Sigma size={18} /> Sort Ngang
                </button>
                <button 
                    onClick={() => setSubTab('cloud')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'cloud' ? '#eff6ff' : 'transparent',
                        color: subTab === 'cloud' ? '#e11d48' : '#64748b', fontWeight: subTab === 'cloud' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Cloud size={18} /> Đồng bộ Cloud
                </button>

                <div style={{ height: '1px', background: '#e2e8f0', margin: '10px 0' }}></div>
                <div style={{ fontSize: '12px', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', marginBottom: '10px' }}>Công cụ Chấm</div>
                <button 
                    onClick={() => setSubTab('math-grading')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'math-grading' ? '#eff6ff' : 'transparent',
                        color: subTab === 'math-grading' ? '#1d4ed8' : '#64748b', fontWeight: subTab === 'math-grading' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Calculator size={18} /> Điểm Toán
                </button>
                <button 
                    onClick={() => setSubTab('science-grading')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'science-grading' ? '#eff6ff' : 'transparent',
                        color: subTab === 'science-grading' ? '#0891b2' : '#64748b', fontWeight: subTab === 'science-grading' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Atom size={18} /> Lý, Hóa, Sinh
                </button>
                <button 
                    onClick={() => setSubTab('it-grading')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'it-grading' ? '#eff6ff' : 'transparent',
                        color: subTab === 'it-grading' ? '#7c3aed' : '#64748b', fontWeight: subTab === 'it-grading' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Monitor size={18} /> Điểm Tin
                </button>
                <button 
                    onClick={() => setSubTab('history-grading')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'history-grading' ? '#eff6ff' : 'transparent',
                        color: subTab === 'history-grading' ? '#c2410c' : '#64748b', fontWeight: subTab === 'history-grading' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <ScrollText size={18} /> Điểm Sử
                </button>
                <button 
                    onClick={() => setSubTab('english-grading')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'english-grading' ? '#eff6ff' : 'transparent',
                        color: subTab === 'english-grading' ? '#15803d' : '#64748b', fontWeight: subTab === 'english-grading' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Globe size={18} /> Điểm Anh
                </button>
            </div>

            <div style={{ flex: 1, padding: '24px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                
                {subTab === 'students' && (
                    <div style={{ display: 'flex', gap: '24px', height: '100%' }}>
                        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                            <div style={{ padding: '16px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <h3 style={{ margin: 0, fontSize: '16px', color: '#1e293b' }}>Danh sách học sinh</h3>
                                <label style={{ 
                                    padding: '8px 16px', background: '#3b82f6', color: 'white', borderRadius: '6px', fontSize: '13px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px'
                                }}>
                                    <Upload size={16} /> Tải file Excel
                                    <input type="file" accept=".xlsx,.xls" hidden onChange={handleStudentUpload} />
                                </label>
                            </div>
                            <div style={{ flex: 1, overflow: 'auto' }}>
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                                    <thead style={{ position: 'sticky', top: 0, background: '#f8fafc', zIndex: 5 }}>
                                        <tr>
                                            <th style={{ padding: '10px', textAlign: 'left', borderBottom: '1px solid #e2e8f0' }}>STT</th>
                                            <th style={{ padding: '10px', textAlign: 'left', borderBottom: '1px solid #e2e8f0' }}>SBD</th>
                                            <th style={{ padding: '10px', textAlign: 'left', borderBottom: '1px solid #e2e8f0' }}>Họ và Tên</th>
                                            <th style={{ padding: '10px', textAlign: 'left', borderBottom: '1px solid #e2e8f0' }}>Lớp</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {students.length > 0 ? students.map((s, idx) => (
                                            <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                                                <td style={{ padding: '10px' }}>{idx + 1}</td>
                                                <td style={{ padding: '10px', fontWeight: 600, color: '#475569' }}>{s.id}</td>
                                                <td style={{ padding: '10px', fontWeight: 500 }}>{s.fullName}</td>
                                                <td style={{ padding: '10px' }}>{s.class}</td>
                                            </tr>
                                        )) : (
                                            <tr>
                                                <td colSpan={4} style={{ padding: '40px', textAlign: 'center', color: '#94a3b8' }}>Chưa có dữ liệu. Vui lòng tải file danh sách.</td>
                                            </tr>
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <div style={{ width: '250px', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', padding: '16px', height: 'fit-content' }}>
                            <h4 style={{ margin: '0 0 15px 0', fontSize: '14px', color: '#475569', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <TrendingUp size={16} /> Thống kê sĩ số
                            </h4>
                            {Object.keys(getClassStats).length > 0 ? (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                    {Object.entries(getClassStats).sort().map(([cls, count]) => (
                                        <div key={cls} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 12px', background: '#f8fafc', borderRadius: '6px', fontSize: '13px' }}>
                                            <span style={{ fontWeight: 600, color: '#1e3a8a' }}>{cls}</span>
                                            <span style={{ fontWeight: 600, color: '#64748b' }}>{count} HS</span>
                                        </div>
                                    ))}
                                    <div style={{ marginTop: '10px', paddingTop: '10px', borderTop: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', fontWeight: 700, fontSize: '13px' }}>
                                        <span>Tổng cộng</span>
                                        <span>{students.length} HS</span>
                                    </div>
                                </div>
                            ) : (
                                <div style={{ fontSize: '12px', color: '#94a3b8', fontStyle: 'italic' }}>Chưa có dữ liệu</div>
                            )}
                        </div>
                    </div>
                )}

                {subTab === 'scores' && (
                     <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                         <div style={{ padding: '10px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', overflowX: 'auto', whiteSpace: 'nowrap', display: 'flex', gap: '8px' }}>
                             {Array.from({length: 40}, (_, i) => i + 1).map(num => (
                                 <button 
                                    key={num}
                                    onClick={() => setActiveExamTime(num)}
                                    style={{
                                        padding: '8px 16px', borderRadius: '6px', border: '1px solid', fontSize: '13px', fontWeight: 600, cursor: 'pointer',
                                        background: activeExamTime === num ? '#1e3a8a' : 'white',
                                        color: activeExamTime === num ? 'white' : '#64748b',
                                        borderColor: activeExamTime === num ? '#1e3a8a' : '#cbd5e1',
                                        minWidth: '70px'
                                    }}
                                 >
                                    Lần {num}
                                 </button>
                             ))}
                         </div>

                         <div style={{ padding: '24px', borderBottom:'1px solid #e2e8f0', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                             <div style={{ marginBottom: '20px', textAlign: 'center' }}>
                                 <h3 style={{ margin: '0 0 10px 0', color: '#1e293b' }}>Dữ liệu điểm - Lần {activeExamTime}</h3>
                                 <p style={{ margin: 0, fontSize: '14px', color: '#64748b' }}>Tải file Excel (.xlsx) chứa sheet "DIEMKHOI". Cột B là SBD, các cột F, G, H, I, J là điểm.</p>
                             </div>

                             <div style={{ display: 'flex', gap: '15px' }}>
                                <label style={{ 
                                        padding: '12px 24px', background: '#22c55e', color: 'white', borderRadius: '8px', fontSize: '14px', fontWeight: 600, 
                                        cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 4px 6px -1px rgba(34, 197, 94, 0.3)'
                                    }}>
                                        <Upload size={18} /> Tải file điểm
                                        <input type="file" accept=".xlsx,.xls,.xlsm" hidden onChange={handleScoreUpload} />
                                </label>
                                
                                {examData[activeExamTime] && Object.keys(examData[activeExamTime]).length > 0 && (
                                    <button 
                                        onClick={handleDeleteScore}
                                        style={{ 
                                            padding: '12px 24px', background: '#ef4444', color: 'white', borderRadius: '8px', fontSize: '14px', fontWeight: 600, 
                                            cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgba(239, 68, 68, 0.3)'
                                        }}>
                                        <Trash2 size={18} /> Xóa dữ liệu
                                    </button>
                                )}
                             </div>
                         </div>

                         <div style={{ flex: 1, overflow: 'auto', background: '#f8fafc', padding: '24px' }}>
                            {activeExamScoreList.length > 0 ? (
                                <div style={{ background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                                    <div style={{ padding: '15px', borderBottom: '1px solid #e2e8f0', fontWeight: 600, color: '#334155', background: '#f1f5f9' }}>
                                        Chi tiết điểm Lần {activeExamTime} ({activeExamScoreList.length} học sinh)
                                    </div>
                                    <div style={{ overflowX: 'auto' }}>
                                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                                            <thead>
                                                <tr>
                                                    <th style={{ padding: '10px', textAlign: 'left' }}>SBD</th>
                                                    <th style={{ padding: '10px', textAlign: 'left' }}>Họ và Tên</th>
                                                    <th style={{ padding: '10px', textAlign: 'left' }}>Lớp</th>
                                                    <th style={{ padding: '10px', textAlign: 'center' }}>Toán</th>
                                                    <th style={{ padding: '10px', textAlign: 'center' }}>Lí</th>
                                                    <th style={{ padding: '10px', textAlign: 'center' }}>Hóa</th>
                                                    <th style={{ padding: '10px', textAlign: 'center' }}>Sinh</th>
                                                    <th style={{ padding: '10px', textAlign: 'center' }}>Anh</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {activeExamScoreList.map((s, idx) => (
                                                    <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9', background: idx % 2 === 0 ? 'white' : '#fcfcfc' }}>
                                                        <td style={{ padding: '10px', fontWeight: 600, color: '#475569' }}>{s.id}</td>
                                                        <td style={{ padding: '10px', fontWeight: 500 }}>{s.fullName}</td>
                                                        <td style={{ padding: '10px' }}>{s.class}</td>
                                                        <td style={{ padding: '10px', textAlign: 'center', color: s.scores.math !== undefined ? '#0f172a' : '#cbd5e1' }}>{s.scores.math ?? '-'}</td>
                                                        <td style={{ padding: '10px', textAlign: 'center', color: s.scores.phys !== undefined ? '#0f172a' : '#cbd5e1' }}>{s.scores.phys ?? '-'}</td>
                                                        <td style={{ padding: '10px', textAlign: 'center', color: s.scores.chem !== undefined ? '#0f172a' : '#cbd5e1' }}>{s.scores.chem ?? '-'}</td>
                                                        <td style={{ padding: '10px', textAlign: 'center', color: s.scores.bio !== undefined ? '#0f172a' : '#cbd5e1' }}>{s.scores.bio ?? '-'}</td>
                                                        <td style={{ padding: '10px', textAlign: 'center', color: s.scores.eng !== undefined ? '#0f172a' : '#cbd5e1' }}>{s.scores.eng ?? '-'}</td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            ) : (
                                <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: '#94a3b8' }}>
                                     <div style={{ width: '60px', height: '60px', borderRadius: '50%', background: '#f1f5f9', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '15px' }}>
                                         <Database size={24} style={{ opacity: 0.3 }} />
                                     </div>
                                     <p>Chưa có dữ liệu điểm cho Lần {activeExamTime}</p>
                                </div>
                            )}
                         </div>
                     </div>
                )}

                {(subTab === 'summary' || subTab === 'sort-summary') && (
                    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                        <div style={{ padding: '12px', borderBottom: '1px solid #e2e8f0', display: 'flex', gap: '8px', background: '#f8fafc', flexWrap: 'wrap', alignItems: 'center' }}>
                             {[
                                 {id: 'math', label: 'Toán'},
                                 {id: 'phys', label: 'Lí'},
                                 {id: 'chem', label: 'Hóa'},
                                 {id: 'eng', label: 'Anh'},
                                 {id: 'bio', label: 'Sinh'},
                                 {id: 'A', label: 'Khối A (T-L-H)'},
                                 {id: 'A1', label: 'Khối A1 (T-L-A)'},
                                 {id: 'B', label: 'Khối B (T-H-S)'},
                                 {id: 'total', label: 'Tổng Khối (Tự động)'},
                             ].map(tab => (
                                 <button 
                                    key={tab.id}
                                    onClick={() => setSummaryTab(tab.id as any)}
                                    style={{
                                        padding: '8px 16px', borderRadius: '6px', border: '1px solid', fontSize: '13px', fontWeight: 600, cursor: 'pointer',
                                        background: summaryTab === tab.id ? '#1e3a8a' : 'white',
                                        color: summaryTab === tab.id ? 'white' : '#475569',
                                        borderColor: summaryTab === tab.id ? '#1e3a8a' : '#cbd5e1',
                                    }}
                                 >
                                    {tab.label}
                                 </button>
                             ))}
                             
                             <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center' }}>
                                {renderMultiSelect()}
                                 <button
                                   onClick={() => exportToExcel('summary-table', `Tong_Ket_${summaryTab}`)}
                                   style={{
                                      padding: '8px 16px', borderRadius: '6px', background: '#059669', border: 'none',
                                      cursor: 'pointer', fontSize: '13px', fontWeight: 600, color: 'white',
                                      display: 'flex', alignItems: 'center', gap: '8px'
                                   }}
                                 >
                                   <FileDown size={14} /> Xuất Excel
                                </button>
                             </div>
                        </div>

                        <div style={{ flex: 1, overflow: 'auto' }}>
                            <table id="summary-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, fontSize: '13px', minWidth: '1200px' }}>
                                <thead style={{ position: 'sticky', top: 0, zIndex: 10, background: '#f1f5f9' }}>
                                    <tr>
                                        <th style={{ padding: '10px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', width: '50px' }}>STT</th>
                                        <th onClick={() => handleSort('id')} style={{ padding: '10px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', cursor: 'pointer', textAlign: 'left' }}>
                                            <div style={{display:'flex', alignItems:'center', gap:'4px'}}>SBD {renderSortIcon('id')}</div>
                                        </th>
                                        <th onClick={() => handleSort('firstName')} style={{ padding: '10px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', cursor: 'pointer', textAlign: 'left', minWidth: '200px' }}>
                                            <div style={{display:'flex', alignItems:'center', gap:'4px'}}>Họ và Tên {renderSortIcon('firstName')}</div>
                                        </th>
                                        <th onClick={() => handleSort('class')} style={{ padding: '10px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', cursor: 'pointer', width: '80px' }}>
                                            <div style={{display:'flex', alignItems:'center', gap:'4px', justifyContent: 'center'}}>Lớp {renderSortIcon('class')}</div>
                                        </th>
                                        {Array.from({length: 40}, (_, i) => i + 1).map(num => (
                                            <th key={num} style={{ padding: '8px', borderBottom: '1px solid #cbd5e1', borderRight: '1px solid #e2e8f0', width: '50px', fontSize: '11px', color: '#64748b' }}>
                                                L{num}
                                            </th>
                                        ))}
                                        <th onClick={() => handleSort('avg')} style={{ padding: '10px', borderBottom: '1px solid #cbd5e1', background: '#e0f2fe', position: 'sticky', right: 0, zIndex: 11, cursor: 'pointer' }}>
                                            <div style={{display:'flex', alignItems:'center', gap:'4px', justifyContent: 'center'}}>TB/Tổng {renderSortIcon('avg')}</div>
                                        </th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sortedData.map((row, idx) => {
                                        let scoresToDisplay: (number|undefined)[] = [];
                                        if (subTab === 'sort-summary') {
                                            const collectedScores: number[] = [];
                                            for (let i = 1; i <= 40; i++) {
                                                const val = row[`score_${i}`];
                                                if (val !== undefined && val !== null) {
                                                    collectedScores.push(val);
                                                }
                                            }
                                            collectedScores.sort((a, b) => b - a);
                                            scoresToDisplay = Array.from({ length: 40 }, (_, i) => collectedScores[i]);
                                        }

                                        return (
                                            <tr key={idx} style={{ background: idx % 2 === 0 ? 'white' : '#fcfcfc' }}>
                                                <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{idx + 1}</td>
                                                <td style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', fontWeight: 600, color: '#475569' }}>{row.id}</td>
                                                <td style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', fontWeight: 500 }}>{row.fullName}</td>
                                                <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{row.class}</td>
                                                
                                                {Array.from({length: 40}, (_, i) => i + 1).map(num => {
                                                    const val = subTab === 'sort-summary'
                                                        ? scoresToDisplay[num - 1]
                                                        : row[`score_${num}`];
                                                    return (
                                                        <td key={num} style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', color: (val !== undefined && val !== null) ? '#0f172a' : '#cbd5e1' }}>
                                                            {(val !== undefined && val !== null) ? val : '-'}
                                                        </td>
                                                    );
                                                })}

                                                <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', background: '#f0f9ff', position: 'sticky', right: 0, fontWeight: 700, color: '#0369a1' }}>
                                                    {row.avg !== null ? row.avg : '-'}
                                                </td>
                                            </tr>
                                        );
                                    })}
                                    {sortedData.length === 0 && (
                                        <tr>
                                            <td colSpan={45} style={{ padding: '40px', textAlign: 'center', color: '#94a3b8' }}>
                                                {students.length === 0 ? "Chưa có dữ liệu. Hãy tải danh sách học sinh và điểm các lần thi." : "Không có học sinh nào phù hợp với bộ lọc này."}
                                            </td>
                                        </tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}
                
                {subTab === 'cloud' && (
                    <div style={{ height: '100%', display: 'flex', gap: '20px' }}>
                        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '20px' }}>
                            <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '16px' }}>
                                    <div style={{ width: '40px', height: '40px', borderRadius: '10px', background: '#ecfdf5', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#059669' }}>
                                        <Share2 size={24} />
                                    </div>
                                    <div>
                                        <h3 style={{ margin: 0, color: '#1e293b' }}>Cấu hình kết nối Google Sheets</h3>
                                        <p style={{ margin: '4px 0 0 0', color: '#64748b', fontSize: '13px' }}>Kết nối để đồng bộ dữ liệu điểm và học sinh lên đám mây.</p>
                                    </div>
                                </div>
                                
                                <div style={{ marginBottom: '16px' }}>
                                    <label style={{ display: 'block', marginBottom: '8px', fontWeight: 600, fontSize: '14px', color: '#475569' }}>
                                        Google Apps Script Web App URL
                                    </label>
                                    <div style={{ display: 'flex', gap: '8px' }}>
                                        <input 
                                            type="text" 
                                            placeholder="https://script.google.com/macros/s/..."
                                            value={scriptUrl}
                                            onChange={(e) => setScriptUrl(e.target.value)}
                                            style={{ flex: 1, padding: '12px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '14px' }}
                                        />
                                        <button 
                                            onClick={() => {
                                                if(!scriptUrl) return;
                                                navigator.clipboard.writeText(scriptUrl);
                                                alert("Đã copy URL vào bộ nhớ đệm!");
                                            }}
                                            title="Copy link để gửi cho người khác"
                                            style={{ padding: '0 16px', borderRadius: '8px', border: '1px solid #cbd5e1', background: '#f8fafc', cursor: 'pointer', color: '#475569' }}
                                        >
                                            <Copy size={18} />
                                        </button>
                                    </div>
                                </div>
                                
                                {/* New Section: Sharing Guide */}
                                <div style={{ padding: '12px', background: '#eff6ff', borderRadius: '8px', border: '1px solid #dbeafe', marginBottom: '16px', fontSize: '13px', color: '#1e40af' }}>
                                    <div style={{ fontWeight: 600, marginBottom: '6px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                        <Laptop size={14} /> Làm sao để đồng bộ với máy khác?
                                    </div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap', justifyContent: 'center', padding: '10px 0' }}>
                                        <div style={{ padding: '6px 12px', background: 'white', borderRadius: '6px', border: '1px solid #bfdbfe', fontWeight: 600 }}>Máy A (Bạn)</div>
                                        <ArrowRight size={14} color="#3b82f6" />
                                        <div style={{ padding: '6px 12px', background: '#dcfce7', borderRadius: '6px', border: '1px solid #86efac', fontWeight: 600, color: '#166534' }}>Google Sheet (Qua URL trên)</div>
                                        <ArrowRight size={14} color="#3b82f6" />
                                        <div style={{ padding: '6px 12px', background: 'white', borderRadius: '6px', border: '1px solid #bfdbfe', fontWeight: 600 }}>Máy B (Người khác)</div>
                                    </div>
                                    <p style={{ margin: '4px 0 0 0' }}>Chỉ cần gửi <b>URL trên</b> cho người khác. Họ dán vào ô bên trên ở máy họ là có thể "Tải dữ liệu" về.</p>
                                    
                                    <div style={{ marginTop: '10px', padding: '8px', background: '#fee2e2', borderRadius: '6px', color: '#991b1b', fontWeight: 600, fontSize: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        <AlertTriangle size={14} /> Lưu ý: URL phải kết thúc bằng "/exec", không dùng link "/edit".
                                    </div>
                                </div>

                                <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
                                     <button 
                                        onClick={handleSyncToCloud}
                                        disabled={syncStatus === 'loading'}
                                        style={{ 
                                            flex: 1, padding: '12px', borderRadius: '8px', background: '#3b82f6', color: 'white', border: 'none', fontWeight: 600, cursor: 'pointer',
                                            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', minWidth: '150px'
                                        }}
                                    >
                                        {syncStatus === 'loading' ? <Loader2 className="animate-spin" size={18}/> : <Upload size={18} />}
                                        Gửi dữ liệu lên Sheet
                                    </button>
                                     <button 
                                        onClick={handleSyncFromCloud}
                                        disabled={syncStatus === 'loading'}
                                        style={{ 
                                            flex: 1, padding: '12px', borderRadius: '8px', background: '#059669', color: 'white', border: 'none', fontWeight: 600, cursor: 'pointer',
                                            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', minWidth: '150px'
                                        }}
                                    >
                                        {syncStatus === 'loading' ? <Loader2 className="animate-spin" size={18}/> : <Download size={18} />}
                                        Tải dữ liệu từ Sheet
                                    </button>
                                </div>

                                {syncMessage && (
                                    <div style={{ marginTop: '16px', padding: '12px', borderRadius: '8px', background: syncStatus === 'error' ? '#fef2f2' : '#f0fdf4', color: syncStatus === 'error' ? '#ef4444' : '#15803d', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        {syncStatus === 'error' ? <X size={16} /> : <CheckSquare size={16} />}
                                        {syncMessage}
                                    </div>
                                )}
                            </div>

                             <div style={{ background: '#f8fafc', padding: '24px', borderRadius: '12px', border: '1px dashed #cbd5e1' }}>
                                <h4 style={{ margin: '0 0 12px 0', color: '#475569', fontSize: '14px' }}>Trạng thái dữ liệu hiện tại</h4>
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                    <div style={{ background: 'white', padding: '12px', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                                        <div style={{ fontSize: '12px', color: '#64748b' }}>Học sinh</div>
                                        <div style={{ fontSize: '20px', fontWeight: 700, color: '#1e3a8a' }}>{students.length}</div>
                                    </div>
                                    <div style={{ background: 'white', padding: '12px', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                                        <div style={{ fontSize: '12px', color: '#64748b' }}>Lần thi có điểm</div>
                                        <div style={{ fontSize: '20px', fontWeight: 700, color: '#1e3a8a' }}>{Object.keys(examData).length}</div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div style={{ width: '400px', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                            <div style={{ padding: '16px', background: '#f1f5f9', borderBottom: '1px solid #e2e8f0', fontWeight: 600, color: '#334155', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <HelpCircle size={18} /> Hướng dẫn cài đặt Script
                            </div>
                            <div style={{ flex: 1, padding: '16px', overflowY: 'auto', fontSize: '13px', lineHeight: '1.5', color: '#475569' }}>
                                <p style={{ marginTop: 0 }}>Để tính năng này hoạt động, bạn cần tạo một Google Apps Script gắn với Google Sheet của bạn:</p>
                                <ol style={{ paddingLeft: '20px' }}>
                                    <li>Mở Google Sheet của bạn.</li>
                                    <li>Vào menu <b>Tiện ích mở rộng</b> &gt; <b>Apps Script</b>.</li>
                                    <li>Xóa code cũ, copy đoạn mã bên dưới và dán vào.</li>
                                    <li>Nhấn nút <b>Triển khai (Deploy)</b> &gt; <b>Tùy chọn triển khai mới</b>.</li>
                                    <li>Chọn loại: <b>Ứng dụng web (Web app)</b>.</li>
                                    <li>Cấu hình:
                                        <ul style={{marginTop: '4px'}}>
                                            <li><b>Thực thi dưới dạng:</b> Tôi (Me)</li>
                                            <li><b>Ai có quyền truy cập:</b> Bất kỳ ai (Anyone)</li>
                                        </ul>
                                    </li>
                                    <li>Nhấn Triển khai, cấp quyền truy cập, sau đó Copy URL nhận được dán vào ô bên trái.</li>
                                </ol>
                                
                                <div style={{ position: 'relative', marginTop: '16px' }}>
                                    <pre style={{ 
                                        background: '#1e293b', color: '#e2e8f0', padding: '12px', borderRadius: '8px', 
                                        overflowX: 'auto', fontSize: '11px', fontFamily: 'monospace', margin: 0, height: '300px'
                                    }}>
                                        {SCRIPT_TEMPLATE}
                                    </pre>
                                    <button 
                                        onClick={() => {
                                            navigator.clipboard.writeText(SCRIPT_TEMPLATE);
                                            alert("Đã copy code vào bộ nhớ đệm!");
                                        }}
                                        style={{ 
                                            position: 'absolute', top: '8px', right: '8px', padding: '6px 12px', 
                                            background: 'rgba(255,255,255,0.1)', color: 'white', border: '1px solid rgba(255,255,255,0.2)', 
                                            borderRadius: '4px', cursor: 'pointer', fontSize: '11px', display: 'flex', gap: '4px', alignItems: 'center'
                                        }}
                                    >
                                        <Copy size={12} /> Copy Code
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                )}
                
                {subTab === 'math-grading' && renderGradingTable(mathResults, 'math', 'Chấm Điểm Toán (40 Câu)', 40)}
                {subTab === 'science-grading' && renderGradingTable(scienceResults, 'science', 'Chấm Điểm KHTN (Lý/Hóa/Sinh)', 40)}
                {subTab === 'it-grading' && renderGradingTable(itResults, 'it', 'Chấm Điểm Tin Học (40 Câu)', 40)}
                {subTab === 'history-grading' && renderGradingTable(historyResults, 'history', 'Chấm Điểm Lịch Sử (40 Câu)', 40)}
                {subTab === 'english-grading' && renderGradingTable(englishResults, 'english', 'Chấm Điểm Tiếng Anh (50 Câu)', 50)}

            </div>
        </div>
    );
};

const App = RankingView;

const root = createRoot(document.getElementById('root')!);
root.render(<App />);