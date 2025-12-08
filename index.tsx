import React, { useState, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI } from "@google/genai";
import { Upload, FileText, Download, Loader2, Settings, Key, Eye, EyeOff, Calculator, FlaskConical, Languages, BrainCircuit, Table as TableIcon, X, User, School, BookOpen, ChevronRight, LayoutDashboard, FileSpreadsheet, RefreshCw, ArrowUpDown, ArrowUp, ArrowDown, FileDown, Filter, Palette, Monitor, Hourglass, TrendingUp, Users, Database, Sigma, Award, Trash2 } from 'lucide-react';

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

interface DocFile {
  id: string;
  name: string;
  content: string; 
  type: 'pdf' | 'text';
}

interface SubjectConfig {
  id: string;
  name: string;
  type: 'math' | 'science' | 'english' | 'it' | 'history';
  totalQuestions: number;
  parts: {
    p1: { start: number; end: number; scorePerQ: number };
    p2: { start: number; end: number; scorePerGroup: number }; 
    p3: { start: number; end: number; scorePerQ: number };
  };
}

interface ThresholdConfig {
  lowCount: number; 
  highPercent: number; 
}

interface ColorConfig {
  lowError: string; 
  highError: string; 
}

// --- Ranking & Summary Types ---

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

// --- Constants ---

const SUBJECTS_CONFIG: Record<string, SubjectConfig> = {
  math: {
    id: 'math',
    name: 'Toán Học',
    type: 'math',
    totalQuestions: 34,
    parts: {
      p1: { start: 1, end: 12, scorePerQ: 0.25 },
      p2: { start: 13, end: 28, scorePerGroup: 1.0 }, 
      p3: { start: 29, end: 34, scorePerQ: 0.5 },
    }
  },
  science: {
    id: 'science',
    name: 'KHTN (Lý/Hóa/Sinh)',
    type: 'science',
    totalQuestions: 40,
    parts: {
      p1: { start: 1, end: 18, scorePerQ: 0.25 },
      p2: { start: 19, end: 34, scorePerGroup: 1.0 }, 
      p3: { start: 35, end: 40, scorePerQ: 0.25 },
    }
  },
  english: {
    id: 'english',
    name: 'Tiếng Anh',
    type: 'english',
    totalQuestions: 40,
    parts: {
      p1: { start: 1, end: 40, scorePerQ: 0.25 },
      p2: { start: 0, end: 0, scorePerGroup: 0 },
      p3: { start: 0, end: 0, scorePerQ: 0 },
    }
  },
  it: {
    id: 'it',
    name: 'Tin học',
    type: 'it',
    totalQuestions: 40, 
    parts: {
      p1: { start: 1, end: 28, scorePerQ: 0.25 },
      p2: { start: 29, end: 40, scorePerGroup: 1.0 },
      p3: { start: 0, end: 0, scorePerQ: 0 },
    }
  },
  history: {
    id: 'history',
    name: 'Lịch sử',
    type: 'history',
    totalQuestions: 40,
    parts: {
      p1: { start: 1, end: 24, scorePerQ: 0.25 },
      p2: { start: 25, end: 40, scorePerGroup: 1.0 }, 
      p3: { start: 0, end: 0, scorePerQ: 0 },
    }
  }
};

const DEFAULT_COLORS = {
    blue: '#dbeafe', 
    yellow: '#fef9c3', 
    red: '#fee2e2'
};

// --- Helper Functions ---

const fileToBase64 = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = reject;
  });
};

const extractTextFromDocx = async (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const arrayBuffer = event.target?.result;
      if (typeof mammoth !== 'undefined') {
        mammoth.extractRawText({ arrayBuffer })
          .then((result: any) => resolve(result.value))
          .catch(reject);
      } else {
        reject("Mammoth not loaded");
      }
    };
    reader.readAsArrayBuffer(file);
  });
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

const getPart2Label = (index: number, subjectType: string): string => {
  const config = SUBJECTS_CONFIG[subjectType];
  if (!config) return String(index);
  const p2 = config.parts.p2;
  
  if (index < p2.start || index > p2.end) return String(index);

  const relativeIndex = index - p2.start;
  const groupNum = Math.floor(relativeIndex / 4) + 1;
  const charCode = 97 + (relativeIndex % 4); // 97 is 'a'
  
  return `${groupNum}${String.fromCharCode(charCode)}`;
};

const exportToExcel = (elementId: string, fileName: string) => {
    const table = document.getElementById(elementId);
    if (!table || typeof XLSX === 'undefined') return;
    const wb = XLSX.utils.table_to_book(table, { sheet: "ThongKe" });
    XLSX.writeFile(wb, `${fileName || 'Thong_ke'}.xlsx`);
};

const exportExamToWord = (content: string, fileName: string) => {
    const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
            <meta charset="utf-8">
            <title>De_On_Tap</title>
            <style>
                body {
                    font-family: 'Times New Roman', serif;
                    font-size: 12pt;
                    line-height: 1.5;
                }
                p { margin-bottom: 10px; }
            </style>
        </head>
        <body>
            ${content.replace(/\n/g, '<br>')}
        </body>
        </html>
    `;

    const blob = new Blob(['\ufeff', htmlContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileName || 'De_On_Tap'}.doc`; 
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};


// --- Scoring Engine ---

const processData = (data: any[], subjectType: 'math' | 'science' | 'english' | 'it' | 'history') => {
  const config = SUBJECTS_CONFIG[subjectType];
  
  const results: StudentResult[] = [];
  const questionStats: Record<number, number> = {}; 
  const correctKeysForDisplay: Record<number, string> = {}; 
  const uniqueKeysPerQuestion: Record<number, Set<string>> = {};

  for (let i = 1; i <= config.totalQuestions; i++) {
    questionStats[i] = 0;
    correctKeysForDisplay[i] = ''; 
    uniqueKeysPerQuestion[i] = new Set();
  }

  data.forEach(row => {
    if (!row['StudentID'] && !row['LastName'] && !row['FirstName']) return;

    const version = String(row['Key Version'] || row['Exam Code'] || row['Mã đề'] || 'default').trim();
    let p1Score = 0;
    let p2Score = 0;
    let p3Score = 0;
    const details: Record<string, 'T' | 'F'> = {};
    const rawAnswers: Record<string, string> = {};

    const checkQuestion = (idx: number) => {
      const stCol = `Stu${idx}`;
      const keyCol = `PriKey${idx}`;
      const stAns = String(row[stCol] || '').trim().toUpperCase();
      let keyAns = String(row[keyCol] || '').trim().toUpperCase();

      if (keyAns) {
          uniqueKeysPerQuestion[idx].add(keyAns);
          if (!correctKeysForDisplay[idx]) correctKeysForDisplay[idx] = keyAns;
      }
      rawAnswers[idx] = stAns;
      if (!keyAns) return false;
      return stAns === keyAns;
    };

    // --- Part 1 ---
    for (let i = config.parts.p1.start; i <= config.parts.p1.end; i++) {
      if (i === 0) continue;
      const isCorrect = checkQuestion(i);
      if (isCorrect) {
        p1Score += config.parts.p1.scorePerQ;
        details[i] = 'T';
      } else {
        details[i] = 'F';
        questionStats[i]++;
      }
    }

    // --- Part 2 ---
    if (config.parts.p2.end > 0) {
      for (let i = config.parts.p2.start; i <= config.parts.p2.end; i += 4) {
        let correctInGroup = 0;
        for (let j = 0; j < 4; j++) {
           const currentQ = i + j;
           if (currentQ > config.parts.p2.end) break;
           const isCorrect = checkQuestion(currentQ);
           if (isCorrect) {
             correctInGroup++;
             details[currentQ] = 'T';
           } else {
             details[currentQ] = 'F';
             questionStats[currentQ]++;
           }
        }
        p2Score += calculateGroupScore(correctInGroup);
      }
    }

    // --- Part 3 ---
    if (config.parts.p3.end > 0) {
      for (let i = config.parts.p3.start; i <= config.parts.p3.end; i++) {
         const isCorrect = checkQuestion(i);
         if (isCorrect) {
           p3Score += config.parts.p3.scorePerQ;
           details[i] = 'T';
         } else {
           details[i] = 'F';
           questionStats[i]++;
         }
      }
    }

    p1Score = Math.round(p1Score * 100) / 100;
    p2Score = Math.round(p2Score * 100) / 100;
    p3Score = Math.round(p3Score * 100) / 100;
    const totalScore = Math.round((p1Score + p2Score + p3Score) * 100) / 100;

    const fName = String(row['FirstName'] || '').trim();
    const lName = String(row['LastName'] || '').trim();
    const fullName = `${fName} ${lName}`.trim(); 

    results.push({
      sbd: String(row['StudentID'] || ''),
      firstName: fName,
      lastName: lName,
      name: fullName,
      code: version,
      rawAnswers,
      scores: { total: totalScore, p1: p1Score, p2: p2Score, p3: p3Score },
      details
    });
  });

  const stats: QuestionStat[] = [];
  const totalStudents = results.length;
  let isMultiVersion = false;
  for (let i = 1; i <= config.totalQuestions; i++) {
      if (uniqueKeysPerQuestion[i].size > 1) {
          isMultiVersion = true;
          break;
      }
  }

  for (let i = 1; i <= config.totalQuestions; i++) {
    stats.push({
      index: i,
      wrongCount: questionStats[i],
      wrongPercent: totalStudents > 0 ? parseFloat(((questionStats[i] / totalStudents) * 100).toFixed(1)) : 0,
      correctKey: isMultiVersion ? '*' : (correctKeysForDisplay[i] || '-')
    });
  }

  return { results, stats };
};

// --- RANKING & SUMMARY COMPONENT ---

const RankingView = () => {
    const [subTab, setSubTab] = useState<'students' | 'scores' | 'summary'>('students');
    const [students, setStudents] = useState<StudentProfile[]>([]);
    const [examData, setExamData] = useState<ExamDataStore>({});
    const [activeExamTime, setActiveExamTime] = useState<number>(1);
    const [summaryTab, setSummaryTab] = useState<'math'|'phys'|'chem'|'eng'|'bio'|'A'|'A1'|'B'|'total'>('math');
    const [sortConfig, setSortConfig] = useState<{ key: string, direction: 'asc'|'desc' } | null>(null);

    const handleStudentUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if(!file) return;

        const reader = new FileReader();
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
                const cl = String(row[3] || '').trim();
                
                parsedStudents.push({
                    id,
                    firstName: firstName,
                    lastName: lastName,
                    fullName: `${lastName} ${firstName}`.trim(),
                    class: cl
                });
            });
            setStudents(parsedStudents);
            e.target.value = ''; 
        };
        reader.readAsBinaryString(file);
    };

    const handleScoreUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if(!file) return;

        const reader = new FileReader();
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
        reader.readAsBinaryString(file);
    };

    const handleDeleteScore = () => {
        if(confirm(`Bạn có chắc muốn đưa toàn bộ điểm Lần ${activeExamTime} về 0 không?`)) {
            setExamData(prev => {
                const newData = { ...prev };
                
                const currentData = newData[activeExamTime];
                if (!currentData) return prev;

                const resetData: Record<string, SubjectScores> = {};
                // Loop through students in this exam set and zero them out
                Object.keys(currentData).forEach(studentId => {
                    resetData[studentId] = {
                        math: 0, 
                        phys: 0, 
                        chem: 0, 
                        bio: 0, 
                        eng: 0
                    };
                });
                newData[activeExamTime] = resetData;
                return newData;
            });
        }
    };

    const getClassStats = useMemo(() => {
        const stats: Record<string, number> = {};
        students.forEach(s => {
            const c = s.class || 'Khác';
            stats[c] = (stats[c] || 0) + 1;
        });
        return stats;
    }, [students]);

    const getComputedData = useMemo(() => {
        if (!students.length) return [];

        return students.map(s => {
            const row: any = { ...s };
            const scores: number[] = [];
            let sum = 0;
            let count = 0;

            for (let i = 1; i <= 40; i++) {
                const record = examData[i]?.[s.id];
                let val: number | undefined = undefined;

                if (record) {
                    if (summaryTab === 'math') val = record.math;
                    else if (summaryTab === 'phys') val = record.phys;
                    else if (summaryTab === 'chem') val = record.chem;
                    else if (summaryTab === 'bio') val = record.bio;
                    else if (summaryTab === 'eng') val = record.eng;
                    else if (summaryTab === 'A') {
                         if (record.math !== undefined && record.phys !== undefined && record.chem !== undefined)
                            val = record.math + record.phys + record.chem;
                    }
                    else if (summaryTab === 'A1') {
                         if (record.math !== undefined && record.phys !== undefined && record.eng !== undefined)
                            val = record.math + record.phys + record.eng;
                    }
                    else if (summaryTab === 'B') {
                         if (record.math !== undefined && record.chem !== undefined && record.bio !== undefined)
                            val = record.math + record.chem + record.bio;
                    }
                    else if (summaryTab === 'total') {
                        const cls = s.class.toUpperCase();
                        let blockSum = 0;
                        let hasData = false;
                        if (cls.includes('E')) {
                             if (record.math !== undefined || record.phys !== undefined || record.eng !== undefined) {
                                blockSum = (record.math || 0) + (record.phys || 0) + (record.eng || 0);
                                hasData = true;
                            }
                        } 
                        else if (cls.includes('B')) {
                             if (record.math !== undefined || record.chem !== undefined || record.bio !== undefined) {
                                blockSum = (record.math || 0) + (record.chem || 0) + (record.bio || 0);
                                hasData = true;
                            }
                        }
                        else if (cls.includes('A')) {
                             if (record.math !== undefined || record.phys !== undefined || record.chem !== undefined) {
                                blockSum = (record.math || 0) + (record.phys || 0) + (record.chem || 0);
                                hasData = true;
                            }
                        }
                        if (hasData) val = blockSum;
                    }
                }

                if (val !== undefined) {
                    row[`score_${i}`] = val;
                    // IGNORE 0 in calculation
                    if (val !== 0) {
                        scores.push(val);
                        sum += val;
                        count++;
                    }
                }
            }

            row.avg = count > 0 ? parseFloat((sum / count).toFixed(2)) : null;
            row.totalVal = sum;
            row.lastScore = scores.length > 0 ? scores[scores.length - 1] : null;

            return row;
        });
    }, [students, examData, summaryTab]);

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

    const activeExamScoreList = useMemo(() => {
        if (!students.length || !examData[activeExamTime]) return [];
        return students.map(s => {
            const sc = examData[activeExamTime][s.id] || {};
            return { ...s, scores: sc };
        }).filter(s => Object.keys(s.scores).length > 0); 
    }, [students, examData, activeExamTime]);

    return (
        <div style={{ display: 'flex', height: '100%', background: '#f8fafc', overflow: 'hidden' }}>
            <div style={{ width: '220px', background: 'white', borderRight: '1px solid #e2e8f0', padding: '20px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <div style={{ fontSize: '12px', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', marginBottom: '10px' }}>Chức năng</div>
                <button 
                    onClick={() => setSubTab('students')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'students' ? '#eff6ff' : 'transparent',
                        color: subTab === 'students' ? '#1e3a8a' : '#64748b',
                        fontWeight: subTab === 'students' ? 600 : 500,
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
                        color: subTab === 'scores' ? '#1e3a8a' : '#64748b',
                        fontWeight: subTab === 'scores' ? 600 : 500,
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
                        color: subTab === 'summary' ? '#1e3a8a' : '#64748b',
                        fontWeight: subTab === 'summary' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <Award size={18} /> Tổng kết
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
                                 <p style={{ margin: 0, fontSize: '14px', color: '#64748b' }}>Tải file Excel (.xlsx) chứa sheet "DIEMKHOI".</p>
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
                                        <Trash2 size={18} /> Xóa dữ liệu (Về 0)
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

                {subTab === 'summary' && (
                    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                        
                        <div style={{ padding: '12px', borderBottom: '1px solid #e2e8f0', display: 'flex', gap: '8px', background: '#f8fafc', flexWrap: 'wrap' }}>
                             {[
                                 {id: 'math', label: 'Toán'},
                                 {id: 'phys', label: 'Lí'},
                                 {id: 'chem', label: 'Hóa'},
                                 {id: 'eng', label: 'Anh'},
                                 {id: 'bio', label: 'Sinh'},
                                 {id: 'A', label: 'Khối A'},
                                 {id: 'A1', label: 'Khối A1'},
                                 {id: 'B', label: 'Khối B'},
                                 {id: 'total', label: 'Tổng Khối'},
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
                             
                             <div style={{ marginLeft: 'auto' }}>
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
                                    {sortedData.map((row, idx) => (
                                        <tr key={idx} style={{ background: idx % 2 === 0 ? 'white' : '#fcfcfc' }}>
                                            <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{idx + 1}</td>
                                            <td style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', fontWeight: 600, color: '#475569' }}>{row.id}</td>
                                            <td style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', fontWeight: 500 }}>{row.fullName}</td>
                                            <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9' }}>{row.class}</td>
                                            
                                            {Array.from({length: 40}, (_, i) => i + 1).map(num => {
                                                const val = row[`score_${num}`];
                                                return (
                                                    <td key={num} style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', borderRight: '1px solid #f1f5f9', color: val ? '#0f172a' : '#cbd5e1' }}>
                                                        {val !== undefined ? val : '-'}
                                                    </td>
                                                );
                                            })}

                                            <td style={{ padding: '8px', textAlign: 'center', borderBottom: '1px solid #f1f5f9', background: '#f0f9ff', position: 'sticky', right: 0, fontWeight: 700, color: '#0369a1' }}>
                                                {row.avg !== null ? row.avg : '-'}
                                            </td>
                                        </tr>
                                    ))}
                                    {sortedData.length === 0 && (
                                        <tr>
                                            <td colSpan={45} style={{ padding: '40px', textAlign: 'center', color: '#94a3b8' }}>
                                                Chưa có dữ liệu. Hãy tải danh sách học sinh và điểm các lần thi.
                                            </td>
                                        </tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}

            </div>
        </div>
    );
};


// --- Main App Component ---

const App = () => {
  const [activeSubject, setActiveSubject] = useState<string>('math');
  const [activeTab, setActiveTab] = useState<'stats' | 'create'>('stats');
  
  const [data, setData] = useState<any[] | null>(null);
  const [processedResults, setProcessedResults] = useState<StudentResult[] | null>(null);
  const [stats, setStats] = useState<QuestionStat[] | null>(null);
  const [fileName, setFileName] = useState<string>("");

  const [statsPartFilter, setStatsPartFilter] = useState<'all' | 'p1' | 'p2' | 'p3'>('all');
  const [questionCounts, setQuestionCounts] = useState<Record<number, number>>({});
  const [sortConfig, setSortConfig] = useState<{ key: string; direction: 'asc' | 'desc' } | null>(null);

  const [thresholds, setThresholds] = useState<ThresholdConfig>(() => {
    const saved = localStorage.getItem('thresholds');
    return saved ? JSON.parse(saved) : { lowCount: 5, highPercent: 40 };
  });

  const [customColors, setCustomColors] = useState<ColorConfig>(() => {
    const saved = localStorage.getItem('customColors');
    return saved ? JSON.parse(saved) : { lowError: DEFAULT_COLORS.yellow, highError: DEFAULT_COLORS.red };
  });

  const [examFile, setExamFile] = useState<DocFile | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedExam, setGeneratedExam] = useState<string>("");

  const [showSettings, setShowSettings] = useState(false);
  const [userApiKey, setUserApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [showKey, setShowKey] = useState(false);

  useEffect(() => {
    localStorage.setItem('gemini_api_key', userApiKey);
  }, [userApiKey]);

  useEffect(() => {
    localStorage.setItem('thresholds', JSON.stringify(thresholds));
  }, [thresholds]);

  useEffect(() => {
    localStorage.setItem('customColors', JSON.stringify(customColors));
  }, [customColors]);

  const handleDataUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    const processBinary = (bstr: string | ArrayBuffer) => {
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const jsonData = XLSX.utils.sheet_to_json(ws);
      setData(jsonData);
      
      const { results, stats } = processData(jsonData, activeSubject as any);
      setProcessedResults(results);
      setStats(stats);
      setSortConfig(null);
    };

    if (file.name.toLowerCase().endsWith('.csv')) {
        reader.onload = (evt) => {
            const text = evt.target?.result as string;
            const wb = XLSX.read(text, { type: 'string' });
            const wsname = wb.SheetNames[0];
            const ws = wb.Sheets[wsname];
            const jsonData = XLSX.utils.sheet_to_json(ws);
            setData(jsonData);
            
            const { results, stats } = processData(jsonData, activeSubject as any);
            setProcessedResults(results);
            setStats(stats);
            setSortConfig(null);
        }
        reader.readAsText(file, 'UTF-8');
    } else {
        reader.onload = (evt) => {
            const bstr = evt.target?.result;
            if (bstr) processBinary(bstr);
        };
        reader.readAsBinaryString(file);
    }
    
    e.target.value = ''; 
  };

  useEffect(() => {
    if (data && activeSubject !== 'ranking') {
      const { results, stats: newStats } = processData(data, activeSubject as any);
      setProcessedResults(results);
      setStats(newStats);
      setQuestionCounts({});
    }
  }, [activeSubject]);

  const handleSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  const sortedResults = useMemo(() => {
    if (!processedResults) return [];
    if (!sortConfig) return processedResults;

    const sorted = [...processedResults];
    sorted.sort((a, b) => {
      let aVal: any = '';
      let bVal: any = '';

      if (sortConfig.key === 'name') {
        aVal = a.firstName; 
        bVal = b.firstName;
      } else if (sortConfig.key === 'sbd') {
        aVal = a.sbd;
        bVal = b.sbd;
      } else if (sortConfig.key === 'total') {
        aVal = a.scores.total;
        bVal = b.scores.total;
      } else if (sortConfig.key === 'p1') {
        aVal = a.scores.p1;
        bVal = b.scores.p1;
      } else if (sortConfig.key === 'p2') {
        aVal = a.scores.p2;
        bVal = b.scores.p2;
      } else if (sortConfig.key === 'p3') {
        aVal = a.scores.p3;
        bVal = b.scores.p3;
      }

      if (aVal < bVal) return sortConfig.direction === 'asc' ? -1 : 1;
      if (aVal > bVal) return sortConfig.direction === 'asc' ? 1 : -1;
      return 0;
    });
    return sorted;
  }, [processedResults, sortConfig]);

  const summaryStats = useMemo(() => {
    if (!processedResults || processedResults.length === 0) return { min: 0, max: 0, avg: 0 };
    const scores = processedResults.map(r => r.scores.total);
    const min = Math.min(...scores);
    const max = Math.max(...scores);
    const sum = scores.reduce((a, b) => a + b, 0);
    const avg = parseFloat((sum / scores.length).toFixed(2));
    return { min, max, avg };
  }, [processedResults]);

  const filteredWrongStats = useMemo(() => {
      if (!stats) return [];
      const config = SUBJECTS_CONFIG[activeSubject];
      if (!config) return [];

      return stats.filter(s => {
          if (s.wrongPercent < thresholds.highPercent) return false;
          if (statsPartFilter === 'all') return true;
          
          const idx = s.index;
          if (statsPartFilter === 'p1') return idx >= config.parts.p1.start && idx <= config.parts.p1.end;
          if (statsPartFilter === 'p2') return idx >= config.parts.p2.start && idx <= config.parts.p2.end;
          if (statsPartFilter === 'p3') return idx >= config.parts.p3.start && idx <= config.parts.p3.end;
          return false;
      }).sort((a,b) => b.wrongCount - a.wrongCount);
  }, [stats, statsPartFilter, thresholds.highPercent, activeSubject]);

  const updateQuestionCount = (index: number, val: number) => {
      setQuestionCounts(prev => ({ ...prev, [index]: val }));
  };

  const handleExamFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    let content = "";
    if (file.type === "application/pdf") {
      content = await fileToBase64(file);
    } else if (file.name.endsWith(".docx") || file.name.endsWith(".doc")) {
      content = await extractTextFromDocx(file);
    } else {
      const reader = new FileReader();
      content = await new Promise((resolve) => {
        reader.onload = (e) => resolve(e.target?.result as string);
        reader.readAsText(file);
      });
    }
    setExamFile({ id: 'exam_orig', name: file.name, content, type: file.type === "application/pdf" ? 'pdf' : 'text' });
  };

  const generatePracticeExam = async () => {
    if (!examFile || !stats) return;
    setIsGenerating(true);

    try {
      const apiKey = userApiKey || process.env.API_KEY || '';
      const ai = new GoogleGenAI({ apiKey: apiKey });
      
      const requestDetails = filteredWrongStats.map(s => {
          const label = getPart2Label(s.index, activeSubject);
          const count = questionCounts[s.index] || 5; 
          return `- Dạng bài câu ${label}: tạo ${count} câu.`;
      }).join('\n');
      
      const prompt = `
        Bạn là một giáo viên chuyên nghiệp. Dưới đây là nội dung của một đề thi gốc và yêu cầu tạo đề ôn tập dựa trên các câu học sinh làm sai nhiều nhất.
        
        Nhiệm vụ:
        Phân tích nội dung kiến thức của các câu hỏi trong đề gốc được liệt kê bên dưới, sau đó tạo nội dung theo cấu trúc sau:

        CẤU TRÚC TRẢ VỀ (Bắt buộc):
        
        1. **Đề gốc**:
           - Trích xuất nội dung các câu hỏi gốc bị sai nhiều từ file đề (Các câu tương ứng với danh sách yêu cầu bên dưới).
        
        2. **Đáp án đề gốc**:
           - Chỉ ghi đáp án trắc nghiệm của các câu gốc đó (Ví dụ: 1.A, 2.C, 3.D...). Không ghi lời giải.
        
        3. **Đề ôn tập**:
           - Tạo các câu hỏi rèn luyện tương tự (đổi số liệu, giữ nguyên dạng bài) cho từng câu sai.
           - Số lượng câu hỏi cho từng dạng bài tuân thủ theo danh sách:
           ${requestDetails}
        
        4. **Đáp án đề ôn tập**:
           - Chỉ ghi đáp án trắc nghiệm của các câu hỏi rèn luyện này (Ví dụ: 1.A, 2.B...). Tuyệt đối không đưa ra lời giải chi tiết.
        
        YÊU CẦU ĐỊNH DẠNG LATEX (TUYỆT ĐỐI TUÂN THỦ):
        1. Tất cả công thức toán học, vật lý, hóa học phải đặt trong cặp dấu $...$.
        2. Ký hiệu Hy Lạp: ρ -> \\rho, θ -> \\theta, α -> \\alpha, β -> \\beta, Δ -> \\Delta, μ -> \\mu, λ -> \\lambda.
        3. Độ C: ◦C -> ^\\circ C. Ví dụ: 300◦C -> $300^\\circ C$, -23◦C -> $-23^\\circ C$.
        4. Phần trăm: % -> \\%.
        5. Đơn vị đo lường: Thêm khoảng cách \\; trước đơn vị.
           - Ví dụ: 50 cm -> $50\\;cm$
           - Ví dụ: 100g -> $100\\;g$
        
        Nội dung đề gốc:
        ${examFile.type === 'text' ? examFile.content : '(Xem PDF đính kèm)'}
      `;

      const parts: any[] = [{ text: prompt }];
      if (examFile.type === 'pdf') {
        parts.push({ inlineData: { mimeType: 'application/pdf', data: examFile.content } });
      }

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: { parts }
      });

      if (response.text) {
        setGeneratedExam(response.text);
      }
    } catch (e: any) {
      alert("Lỗi AI: " + e.message);
    } finally {
      setIsGenerating(false);
    }
  };

  const getCellColor = (wrongCount: number, wrongPercent: number) => {
    if (wrongCount === 0) return DEFAULT_COLORS.blue; 
    if (wrongPercent > thresholds.highPercent) return customColors.highError; 
    if (wrongCount < thresholds.lowCount) return customColors.lowError; 
    return 'white';
  };

  const getScoreColor = (score: number) => {
    if (score >= 8) return '#16a34a'; 
    if (score >= 5) return '#ca8a04'; 
    return '#dc2626'; 
  };

  const renderSortIcon = (key: string) => {
    if (!sortConfig || sortConfig.key !== key) return <ArrowUpDown size={12} style={{ opacity: 0.3 }} />;
    return sortConfig.direction === 'asc' ? <ArrowUp size={12} /> : <ArrowDown size={12} />;
  };

  if (showSettings) {
    return (
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(15, 23, 42, 0.6)', backdropFilter: 'blur(4px)', zIndex: 2000, display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
         <div style={{ background: 'white', borderRadius: '24px', boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)', width: '600px', overflow: 'hidden' }}>
             <div style={{ background: '#1e3a8a', padding: '20px 30px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                 <h2 style={{ margin: 0, color: 'white', fontSize: '18px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <Settings size={20} /> Cài đặt hệ thống
                 </h2>
                 <button onClick={() => setShowSettings(false)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: 'white', borderRadius: '50%', padding: '6px', cursor: 'pointer', display: 'flex' }}>
                    <X size={18} />
                 </button>
             </div>

             <div style={{ padding: '30px', maxHeight: '70vh', overflowY: 'auto' }}>
                 
                 <div style={{ marginBottom: '30px', background: '#eff6ff', padding: '20px', borderRadius: '16px', display: 'flex', gap: '20px', alignItems: 'center' }}>
                     <div style={{ width: '64px', height: '64px', borderRadius: '50%', background: '#3b82f6', color: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '24px', fontWeight: 'bold' }}>H</div>
                     <div>
                         <h3 style={{ margin: '0 0 5px 0', color: '#1e3a8a', fontSize: '18px' }}>Nguyễn Đức Hiền</h3>
                         <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#1d4ed8', fontSize: '14px', marginBottom: '4px' }}>
                             <User size={14} /> Giáo viên Vật Lí
                         </div>
                         <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#64748b', fontSize: '13px' }}>
                             <School size={14} /> Trường THCS và THPT Nguyễn Khuyến Bình Dương
                         </div>
                     </div>
                 </div>

                 <div style={{ marginBottom: '30px' }}>
                    <h4 style={{ margin: '0 0 15px 0', color: '#334155', display: 'flex', alignItems: 'center', gap: '8px' }}>
                         <Filter size={16} /> Cấu hình hiển thị thống kê
                    </h4>
                    
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '20px' }}>
                       <div>
                          <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                             Ngưỡng sai ít (Số lượng)
                          </label>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                             <span style={{ fontSize: '13px', color: '#64748b' }}>&lt;</span>
                             <input 
                                type="number" 
                                value={thresholds.lowCount}
                                onChange={(e) => setThresholds({...thresholds, lowCount: parseInt(e.target.value) || 0})}
                                style={{ width: '60px', padding: '8px', borderRadius: '6px', border: '1px solid #cbd5e1' }}
                             />
                             <span style={{ fontSize: '13px', color: '#64748b' }}>học sinh</span>
                          </div>
                       </div>
                       <div>
                          <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                             Ngưỡng sai nhiều (Tỷ lệ %)
                          </label>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                             <span style={{ fontSize: '13px', color: '#64748b' }}>&gt;</span>
                             <input 
                                type="number" 
                                value={thresholds.highPercent}
                                onChange={(e) => setThresholds({...thresholds, highPercent: parseInt(e.target.value) || 0})}
                                style={{ width: '60px', padding: '8px', borderRadius: '6px', border: '1px solid #cbd5e1' }}
                             />
                             <span style={{ fontSize: '13px', color: '#64748b' }}>%</span>
                          </div>
                       </div>
                    </div>

                    <h4 style={{ margin: '0 0 15px 0', color: '#334155', display: 'flex', alignItems: 'center', gap: '8px' }}>
                         <Palette size={16} /> Màu sắc hiển thị
                    </h4>
                    
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                        <div>
                           <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                              Màu báo sai ít (Mặc định: Vàng)
                           </label>
                           <div style={{display:'flex', gap:'10px', alignItems:'center'}}>
                               <input 
                                  type="color" 
                                  value={customColors.lowError}
                                  onChange={(e) => setCustomColors({...customColors, lowError: e.target.value})}
                                  style={{ height: '36px', width: '60px', padding: '0', border: 'none', cursor: 'pointer', borderRadius: '4px' }}
                               />
                               <span style={{fontSize:'12px', color:'#64748b', fontFamily:'monospace'}}>{customColors.lowError}</span>
                           </div>
                        </div>
                        <div>
                           <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                              Màu báo sai nhiều (Mặc định: Đỏ)
                           </label>
                           <div style={{display:'flex', gap:'10px', alignItems:'center'}}>
                               <input 
                                  type="color" 
                                  value={customColors.highError}
                                  onChange={(e) => setCustomColors({...customColors, highError: e.target.value})}
                                  style={{ height: '36px', width: '60px', padding: '0', border: 'none', cursor: 'pointer', borderRadius: '4px' }}
                               />
                               <span style={{fontSize:'12px', color:'#64748b', fontFamily:'monospace'}}>{customColors.highError}</span>
                           </div>
                        </div>
                    </div>
                 </div>

                 <div>
                     <label style={{ display: 'block', marginBottom: '10px', fontWeight: 600, color: '#334155', fontSize: '15px' }}>
                         Google Gemini API Key
                     </label>
                     <div style={{ position: 'relative' }}>
                         <input 
                             type={showKey ? "text" : "password"} 
                             value={userApiKey}
                             onChange={(e) => setUserApiKey(e.target.value)}
                             placeholder="Nhập API Key của bạn tại đây..."
                             style={{ 
                                 width: '100%', padding: '14px 50px 14px 16px', borderRadius: '12px', border: '1px solid #cbd5e1', 
                                 fontSize: '16px', outline: 'none', transition: 'border 0.2s', background: 'white', color: '#1e293b'
                             }}
                             onFocus={(e) => e.target.style.borderColor = '#2563eb'}
                             onBlur={(e) => e.target.style.borderColor = '#cbd5e1'}
                         />
                         <button 
                            onClick={() => setShowKey(!showKey)} 
                            style={{ 
                                position: 'absolute', right: '12px', top: '50%', transform: 'translateY(-50%)', 
                                background: 'transparent', border: 'none', cursor: 'pointer', color: '#94a3b8', padding: '5px' 
                            }}>
                            {showKey ? <EyeOff size={20}/> : <Eye size={20}/>}
                         </button>
                     </div>
                     <p style={{ marginTop: '10px', fontSize: '13px', color: '#64748b' }}>
                        API Key sẽ được lưu trên trình duyệt của bạn để sử dụng cho các lần sau.
                     </p>
                 </div>

             </div>

             <div style={{ padding: '20px 30px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', justifyContent: 'flex-end' }}>
                 <button 
                    onClick={() => setShowSettings(false)} 
                    style={{ 
                        padding: '10px 24px', background: '#1e3a8a', color: 'white', borderRadius: '8px', border: 'none', 
                        cursor: 'pointer', fontWeight: 600, fontSize: '14px', boxShadow: '0 4px 6px -1px rgba(30, 58, 138, 0.2)'
                    }}>
                    Đóng
                 </button>
             </div>
         </div>
      </div>
    );
  }

  const p2Range = SUBJECTS_CONFIG[activeSubject] ? SUBJECTS_CONFIG[activeSubject].parts.p2 : { start: 0, end: 0 };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', overflow: 'hidden', background: '#f8fafc' }}>
      
      <header style={{ height: '64px', background: '#1e3a8a', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 24px', color: 'white', flexShrink: 0 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', fontWeight: 700, fontSize: '18px' }}>
              <div style={{ width: '36px', height: '36px', background: 'white', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#1e3a8a' }}>
                  <FileSpreadsheet size={20} />
              </div>
              <div>
                  <div style={{ lineHeight: '1.2' }}>THỐNG KÊ CÂU SAI</div>
                  <div style={{ fontSize: '16px', color: '#93c5fd', fontWeight: 500 }}>NK12</div>
              </div>
          </div>
          
          <button 
              onClick={() => setShowSettings(true)}
              style={{ background: 'rgba(255,255,255,0.1)', border: 'none', color: 'white', borderRadius: '8px', padding: '8px', cursor: 'pointer', display: 'flex' }}
          >
              <Settings size={20} />
          </button>
      </header>

      <div style={{ background: 'white', borderBottom: '1px solid #e2e8f0', padding: '12px 24px', display: 'flex', gap: '10px', overflowX: 'auto' }}>
          {Object.values(SUBJECTS_CONFIG).map(subj => {
              const isActive = activeSubject === subj.id;
              return (
                <button
                  key={subj.id}
                  onClick={() => setActiveSubject(subj.type)}
                  style={{
                      display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 20px', 
                      borderRadius: '99px', 
                      border: 'none', cursor: 'pointer', 
                      background: isActive ? '#1e3a8a' : '#f1f5f9',
                      color: isActive ? 'white' : '#64748b',
                      fontWeight: isActive ? 600 : 500,
                      transition: 'all 0.2s ease',
                      fontSize: '14px',
                      boxShadow: isActive ? '0 4px 6px -1px rgba(30, 58, 138, 0.2)' : 'none'
                  }}
                >
                  {subj.id === 'math' && <Calculator size={18} />}
                  {subj.id === 'science' && <FlaskConical size={18} />}
                  {subj.id === 'english' && <Languages size={18} />}
                  {subj.id === 'it' && <Monitor size={18} />}
                  {subj.id === 'history' && <Hourglass size={18} />}
                  <span>{subj.name}</span>
                </button>
              );
          })}
          
          <button
            onClick={() => setActiveSubject('ranking')}
            style={{
                display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 20px', 
                borderRadius: '99px',
                border: 'none', cursor: 'pointer', 
                background: activeSubject === 'ranking' ? '#1e3a8a' : '#f1f5f9',
                color: activeSubject === 'ranking' ? 'white' : '#64748b',
                fontWeight: activeSubject === 'ranking' ? 600 : 500,
                transition: 'all 0.2s ease',
                fontSize: '14px',
                boxShadow: activeSubject === 'ranking' ? '0 4px 6px -1px rgba(30, 58, 138, 0.2)' : 'none'
            }}
          >
            <Award size={18} />
            <span>Tổng kết và Xếp hạng</span>
          </button>
      </div>

      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          
          {activeSubject !== 'ranking' && (
              <div style={{ padding: '15px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', gap: '8px' }}>
                      <button 
                        onClick={() => setActiveTab('stats')}
                        style={{ 
                            padding: '10px 20px', borderRadius: '99px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 600,
                            background: activeTab === 'stats' ? '#1e3a8a' : '#e2e8f0',
                            color: activeTab === 'stats' ? 'white' : '#64748b',
                            boxShadow: activeTab === 'stats' ? '0 4px 6px -1px rgba(30, 58, 138, 0.2)' : 'none',
                            display: 'flex', alignItems: 'center', gap: '8px', transition: 'all 0.2s'
                        }}>
                        <TableIcon size={16} /> Thống kê điểm
                      </button>
                      <button 
                        onClick={() => setActiveTab('create')}
                        style={{ 
                            padding: '10px 20px', borderRadius: '99px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 600,
                            background: activeTab === 'create' ? '#1e3a8a' : '#e2e8f0',
                            color: activeTab === 'create' ? 'white' : '#64748b',
                            boxShadow: activeTab === 'create' ? '0 4px 6px -1px rgba(30, 58, 138, 0.2)' : 'none',
                            display: 'flex', alignItems: 'center', gap: '8px', transition: 'all 0.2s'
                        }}>
                        <BrainCircuit size={16} /> Phân tích AI
                      </button>
                  </div>

                  {activeTab === 'stats' && (
                      <div style={{ display: 'flex', gap: '10px' }}>
                         {stats && (
                            <button
                               onClick={() => exportToExcel('stats-table', fileName)}
                               style={{
                                  padding: '10px 20px', borderRadius: '99px', background: '#2563eb', border: 'none',
                                  cursor: 'pointer', fontSize: '14px', fontWeight: 600, color: 'white',
                                  display: 'flex', alignItems: 'center', gap: '8px',
                                  boxShadow: '0 2px 4px rgba(37, 99, 235, 0.2)'
                               }}
                            >
                               <FileDown size={16} /> Tải file Excel
                            </button>
                         )}
                         <button
                            onClick={() => document.getElementById('re-upload')?.click()}
                            style={{
                               padding: '10px 20px', borderRadius: '99px', background: 'white', border: '1px solid #cbd5e1',
                               cursor: 'pointer', fontSize: '14px', fontWeight: 600, color: '#475569',
                               display: 'flex', alignItems: 'center', gap: '8px',
                               boxShadow: '0 1px 2px rgba(0,0,0,0.05)'
                            }}
                         >
                            <RefreshCw size={16} /> Tải file khác
                         </button>
                         <input type="file" accept=".xlsx,.csv" onChange={handleDataUpload} style={{ display: 'none' }} id="re-upload" />
                      </div>
                  )}
              </div>
          )}

          <div style={{ flex: 1, overflow: 'auto', padding: activeSubject === 'ranking' ? 0 : '0 24px 24px 24px' }}>
             
                {activeSubject === 'ranking' ? (
                    <RankingView />
                ) : (
                    <>
                    {activeTab === 'stats' && (
                        <div style={{ animation: 'fadeIn 0.3s ease-out', height: '100%', display: 'flex', flexDirection: 'column' }}>
                            {!stats && (
                               <div style={{ 
                                    padding: '60px', background: 'white', borderRadius: '16px', border: '2px dashed #cbd5e1', 
                                    textAlign: 'center', transition: 'border-color 0.2s', maxWidth: '600px', margin: '40px auto'
                                }}
                                onDragOver={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#3b82f6'; }}
                                onDragLeave={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#cbd5e1'; }}
                                onDrop={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#cbd5e1'; }}
                                >
                                    <input type="file" accept=".xlsx,.csv" onChange={handleDataUpload} style={{ display: 'none' }} id="data-upload" />
                                    <label htmlFor="data-upload" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '15px' }}>
                                        <div style={{ width: '80px', height: '80px', background: '#eff6ff', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                            <Upload size={36} color="#3b82f6" />
                                        </div>
                                        <div>
                                            <div style={{ fontSize: '18px', fontWeight: 600, color: '#1e293b' }}>
                                                Tải lên file ZipGrade
                                            </div>
                                            <div style={{ fontSize: '14px', color: '#64748b', marginTop: '6px' }}>Hỗ trợ định dạng Excel (.xlsx) hoặc CSV</div>
                                        </div>
                                    </label>
                                </div>
                            )}

                            {processedResults && stats && (
                                <div style={{ background: 'white', borderRadius: '12px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', display: 'flex', flexDirection: 'column', height: '100%', border: '1px solid #e2e8f0' }}>
                                    <div style={{ padding: '12px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTopLeftRadius: '12px', borderTopRightRadius: '12px' }}>
                                        <div style={{ display: 'flex', gap: '24px', fontSize: '13px', fontWeight: 500, alignItems: 'center' }}>
                                            <div style={{ display: 'flex', gap: '15px', borderRight: '1px solid #cbd5e1', paddingRight: '15px' }}>
                                                <div style={{ 
                                                    display: 'flex', gap: '6px', alignItems: 'center', padding: '4px 8px', borderRadius: '6px', 
                                                    background: 'white', border: '1px solid #e2e8f0', color: '#1e3a8a', fontWeight: 600
                                                }}>
                                                    <span style={{width:'12px', height:'12px', background: DEFAULT_COLORS.blue, borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> 
                                                    0 Sai
                                                </div>
                                                <div style={{ 
                                                    display: 'flex', gap: '6px', alignItems: 'center', padding: '4px 8px', borderRadius: '6px', 
                                                    background: 'white', border: '1px solid #e2e8f0', color: '#1e3a8a', fontWeight: 600
                                                }}>
                                                    <span style={{width:'12px', height:'12px', background: customColors.lowError, borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> 
                                                    &lt;{thresholds.lowCount} Sai
                                                </div>
                                                <div style={{ 
                                                    display: 'flex', gap: '6px', alignItems: 'center', padding: '4px 8px', borderRadius: '6px', 
                                                    background: 'white', border: '1px solid #e2e8f0', color: '#1e3a8a', fontWeight: 600
                                                }}>
                                                    <span style={{width:'12px', height:'12px', background: customColors.highError, borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> 
                                                    &gt;{thresholds.highPercent}% Sai
                                                </div>
                                            </div>
                                            <div style={{ display: 'flex', gap: '15px', color: '#334155' }}>
                                                <div>TB: <strong>{summaryStats.avg}</strong></div>
                                                <div>Max: <strong style={{color:'#16a34a'}}>{summaryStats.max}</strong></div>
                                                <div>Min: <strong style={{color:'#dc2626'}}>{summaryStats.min}</strong></div>
                                            </div>
                                        </div>
                                        <div style={{ fontSize: '13px', color: '#64748b' }}>
                                            File: <strong>{fileName}</strong>
                                        </div>
                                    </div>

                                    <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
                                        <table id="stats-table" style={{ width: '100%', fontSize: '12px', borderCollapse: 'separate', borderSpacing: 0, minWidth: '1500px' }}>
                                            <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                                                <tr style={{ background: '#f1f5f9' }}>
                                                    <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', width: '40px' }}>STT</th>
                                                    <th onClick={() => handleSort('sbd')} style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', cursor: 'pointer', userSelect: 'none' }}>
                                                       <div style={{display:'flex', alignItems:'center', justifyContent:'center', gap:'4px'}}>SBD {renderSortIcon('sbd')}</div>
                                                    </th>
                                                    <th onClick={() => handleSort('name')} style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', textAlign: 'left', paddingLeft: '10px', cursor: 'pointer', userSelect: 'none', minWidth: '220px' }}>
                                                       <div style={{display:'flex', alignItems:'center', gap:'4px'}}>Họ và Tên {renderSortIcon('name')}</div>
                                                    </th>
                                                    <th style={{ borderBottom: '1px solid #cbd5e1' }}>Mã</th>
                                                    <th onClick={() => handleSort('total')} style={{ borderBottom: '1px solid #cbd5e1', cursor: 'pointer', userSelect: 'none' }}>
                                                       <div style={{display:'flex', alignItems:'center', justifyContent:'center', gap:'4px'}}>Điểm {renderSortIcon('total')}</div>
                                                    </th>
                                                    <th onClick={() => handleSort('p1')} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '10px', cursor: 'pointer', userSelect: 'none' }}>P1 {renderSortIcon('p1')}</th>
                                                    <th onClick={() => handleSort('p2')} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '10px', cursor: 'pointer', userSelect: 'none' }}>P2 {renderSortIcon('p2')}</th>
                                                    <th onClick={() => handleSort('p3')} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '10px', borderRight: '2px solid #94a3b8', cursor: 'pointer', userSelect: 'none' }}>P3 {renderSortIcon('p3')}</th>
                                                    {stats.map(s => {
                                                       const isPart2 = p2Range && s.index >= p2Range.start && s.index <= p2Range.end;
                                                       return (
                                                        <th key={s.index} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '11px', minWidth: '24px', background: isPart2 ? '#fefce8' : '#f1f5f9' }}>
                                                            {getPart2Label(s.index, activeSubject as any)}
                                                        </th>
                                                       );
                                                    })}
                                                </tr>
                                                
                                                <tr style={{ background: '#e2e8f0' }}>
                                                    <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#e2e8f0', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1' }}></th>
                                                    <th style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#e2e8f0', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1' }}></th>
                                                    <th style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#e2e8f0', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', textAlign: 'left', paddingLeft: '10px', color: '#475569', fontSize: '11px', minWidth: '220px' }}>Đáp án đúng</th>
                                                    <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                    <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                    <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                    <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                    <th style={{ borderBottom: '1px solid #cbd5e1', borderRight: '2px solid #94a3b8' }}></th>
                                                    {stats.map(s => {
                                                       const isPart2 = p2Range && s.index >= p2Range.start && s.index <= p2Range.end;
                                                       return (
                                                        <th key={s.index} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '11px', color: '#16a34a', background: isPart2 ? '#fefce8' : '#e2e8f0' }}>
                                                            {s.correctKey}
                                                        </th>
                                                       );
                                                    })}
                                                </tr>

                                                <tr style={{ background: '#f8fafc' }}>
                                                    <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#f8fafc', borderRight: '1px solid #cbd5e1', borderBottom: '2px solid #94a3b8' }}></th>
                                                    <th style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#f8fafc', borderRight: '1px solid #cbd5e1', borderBottom: '2px solid #94a3b8' }}></th>
                                                    <th style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#f8fafc', borderRight: '1px solid #cbd5e1', borderBottom: '2px solid #94a3b8', textAlign: 'left', paddingLeft: '10px', color: '#64748b', fontSize: '11px', minWidth: '220px' }}>Thống kê (Số lượng/%)</th>
                                                    <th style={{ borderBottom: '2px solid #94a3b8' }}></th>
                                                    <th style={{ borderBottom: '2px solid #94a3b8' }}></th>
                                                    <th style={{ borderBottom: '2px solid #94a3b8' }}></th>
                                                    <th style={{ borderBottom: '2px solid #94a3b8' }}></th>
                                                    <th style={{ borderBottom: '2px solid #94a3b8', borderRight: '2px solid #94a3b8' }}></th>
                                                    {stats.map(s => (
                                                        <th key={s.index} style={{ padding: '4px', background: getCellColor(s.wrongCount, s.wrongPercent), borderBottom: '2px solid #94a3b8', minWidth: '24px', fontSize: '10px', color: '#475569' }}>
                                                            <div>{s.wrongCount}</div>
                                                            <div style={{fontSize: '9px', opacity: 0.8}}>{s.wrongPercent}%</div>
                                                        </th>
                                                    ))}
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {sortedResults.map((st, idx) => (
                                                    <tr key={idx} style={{ background: idx % 2 === 0 ? 'white' : '#fcfcfc' }}>
                                                        <td style={{ position: 'sticky', left: 0, background: idx % 2 === 0 ? 'white' : '#fcfcfc', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9' }}>{idx + 1}</td>
                                                        <td style={{ position: 'sticky', left: '40px', background: idx % 2 === 0 ? 'white' : '#fcfcfc', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9', fontFamily: 'monospace' }}>{st.sbd}</td>
                                                        <td style={{ 
                                                            position: 'sticky', left: '100px', background: idx % 2 === 0 ? 'white' : '#fcfcfc', 
                                                            borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9', 
                                                            textAlign: 'left', fontWeight: 600, paddingLeft: '10px', verticalAlign: 'middle', minWidth: '220px'
                                                        }}>
                                                            <div style={{ 
                                                                display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', 
                                                                overflow: 'hidden', textOverflow: 'ellipsis', lineHeight: '1.4', maxHeight: '2.8em' 
                                                            }}>
                                                                {st.name}
                                                            </div>
                                                        </td>
                                                        <td style={{ borderBottom: '1px solid #f1f5f9' }}>{st.code}</td>
                                                        <td style={{ borderBottom: '1px solid #f1f5f9', fontWeight: 'bold', color: getScoreColor(st.scores.total) }}>{st.scores.total}</td>
                                                        <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px' }}>{st.scores.p1}</td>
                                                        <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px' }}>{st.scores.p2}</td>
                                                        <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px', borderRight: '2px solid #e2e8f0' }}>{st.scores.p3}</td>
                                                        {stats.map(s => {
                                                            const isCorrect = st.details[s.index] === 'T';
                                                            const isPart2 = p2Range && s.index >= p2Range.start && s.index <= p2Range.end;
                                                            let bgColor = 'transparent';
                                                            if (isCorrect) {
                                                                bgColor = isPart2 ? DEFAULT_COLORS.yellow : 'transparent';
                                                            } else {
                                                                bgColor = '#fecaca'; 
                                                            }
                                                            return (
                                                                <td key={s.index} style={{ 
                                                                    borderBottom: '1px solid #f1f5f9', 
                                                                    background: bgColor,
                                                                    color: isCorrect ? '#cbd5e1' : '#b91c1c',
                                                                    fontSize: '11px', fontWeight: isCorrect ? 400 : 700,
                                                                    padding: '2px',
                                                                    borderLeft: '1px solid #f1f5f9'
                                                                }}>
                                                                    {isCorrect ? '•' : st.rawAnswers[s.index]}
                                                                </td>
                                                            );
                                                        })}
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            )}
                        </div>
                    )}

                    {activeTab === 'create' && (
                        <div style={{ display: 'grid', gridTemplateColumns: '350px 1fr', gap: '30px', animation: 'fadeIn 0.3s ease-out', maxWidth: '1400px', margin: '0 auto', height: '100%' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
                            <div style={{ background: 'white', padding: '20px', borderRadius: '16px', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', height: '100%' }}>
                                <h3 style={{ margin: '0 0 15px 0', fontSize: '16px', color: '#1e3a8a', display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
                                    <BookOpen size={18} /> Cấu hình tạo đề
                                </h3>
                                
                                <div style={{ 
                                    padding: '20px', background: '#f8fafc', borderRadius: '12px', border: '2px dashed #e2e8f0', 
                                    textAlign: 'center', marginBottom: '15px', cursor: 'pointer', transition: 'all 0.2s', flexShrink: 0
                                }}
                                onClick={() => document.getElementById('exam-upload')?.click()}
                                onMouseOver={(e) => e.currentTarget.style.borderColor = '#93c5fd'}
                                onMouseOut={(e) => e.currentTarget.style.borderColor = '#e2e8f0'}
                                >
                                    <input type="file" accept=".pdf,.docx,.doc" onChange={handleExamFileUpload} style={{ display: 'none' }} id="exam-upload" />
                                    <div style={{ width: '36px', height: '36px', background: 'white', borderRadius: '50%', margin: '0 auto 8px auto', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                                        <Upload size={18} color="#64748b" />
                                    </div>
                                    <div style={{ fontSize: '13px', fontWeight: 600, color: '#334155' }}>
                                        {examFile ? examFile.name : "Chọn file đề gốc"}
                                    </div>
                                    <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '4px' }}>Hỗ trợ PDF, DOCX</div>
                                </div>

                                {stats && (
                                    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0, marginBottom: '15px' }}>
                                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px', flexShrink: 0 }}>
                                            <div style={{ fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                                                Các câu sai nhiều (&gt;{thresholds.highPercent}%):
                                            </div>
                                            <select 
                                                value={statsPartFilter} 
                                                onChange={(e) => setStatsPartFilter(e.target.value as any)}
                                                style={{ 
                                                    padding: '4px 8px', borderRadius: '6px', 
                                                    border: '1px solid #cbd5e1', background: 'white',
                                                    fontSize: '11px', color: '#1e3a8a', fontWeight: 600,
                                                    cursor: 'pointer', outline: 'none' 
                                                }}
                                            >
                                                <option value="all">Tất cả</option>
                                                <option value="p1">Phần 1</option>
                                                <option value="p2">Phần 2</option>
                                                <option value="p3">Phần 3</option>
                                            </select>
                                        </div>
                                        
                                        <div style={{ overflowY: 'auto', flex: 1, paddingRight: '4px' }}>
                                            {filteredWrongStats.length > 0 ? (
                                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                                    {filteredWrongStats.map(s => (
                                                        <div key={s.index} style={{ 
                                                            display: 'flex', alignItems: 'center', justifyContent: 'space-between', 
                                                            padding: '8px 10px', background: '#fff1f2', border: '1px solid #fecaca', borderRadius: '8px' 
                                                        }}>
                                                            <div style={{ display: 'flex', flexDirection: 'column' }}>
                                                                <span style={{ fontSize: '13px', fontWeight: 700, color: '#be123c' }}>
                                                                    Câu {getPart2Label(s.index, activeSubject as any)}
                                                                </span>
                                                                <span style={{ fontSize: '11px', color: '#881337' }}>
                                                                    Sai: {s.wrongPercent}% ({s.wrongCount} HS)
                                                                </span>
                                                            </div>
                                                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                                                <span style={{ fontSize: '11px', color: '#475569' }}>Số câu:</span>
                                                                <input 
                                                                    type="number" 
                                                                    min="1" 
                                                                    max="50"
                                                                    value={questionCounts[s.index] || 5} 
                                                                    onChange={(e) => updateQuestionCount(s.index, parseInt(e.target.value) || 0)}
                                                                    style={{ 
                                                                        width: '50px', padding: '4px', borderRadius: '4px', border: '1px solid #cbd5e1', 
                                                                        textAlign: 'center', fontSize: '13px', fontWeight: 600
                                                                    }}
                                                                />
                                                            </div>
                                                        </div>
                                                    ))}
                                                </div>
                                            ) : (
                                                <div style={{ fontSize: '12px', color: '#94a3b8', fontStyle: 'italic', padding: '10px', textAlign: 'center' }}>
                                                    Không có câu nào thỏa mãn điều kiện lọc.
                                                </div>
                                            )}
                                        </div>
                                    </div>
                                )}

                                <button 
                                    onClick={generatePracticeExam}
                                    disabled={!examFile || !stats || isGenerating}
                                    style={{ 
                                        width: '100%', padding: '12px', background: (!examFile || !stats) ? '#cbd5e1' : '#1e3a8a', 
                                        color: 'white', border: 'none', borderRadius: '10px', cursor: (!examFile || !stats) ? 'not-allowed' : 'pointer',
                                        fontWeight: 600, display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '10px',
                                        boxShadow: (!examFile || !stats) ? 'none' : '0 4px 6px -1px rgba(30, 58, 138, 0.3)',
                                        transition: 'transform 0.1s', flexShrink: 0
                                    }}
                                    onMouseDown={(e) => !isGenerating && (e.currentTarget.style.transform = 'scale(0.98)')}
                                    onMouseUp={(e) => !isGenerating && (e.currentTarget.style.transform = 'scale(1)')}
                                >
                                    {isGenerating ? <Loader2 className="spin" size={20} /> : <BrainCircuit size={20} />} 
                                    {isGenerating ? 'Đang phân tích...' : 'Tạo đề ôn tập'}
                                </button>
                            </div>
                        </div>

                        <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                            <div style={{ flex: 1, background: 'white', borderRadius: '16px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', overflow: 'hidden', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)' }}>
                                <div style={{ padding: '15px 20px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <div style={{ fontWeight: 600, color: '#334155', fontSize: '14px' }}>Nội dung đề tạo bởi AI</div>
                                    {generatedExam && (
                                        <button 
                                            onClick={() => exportExamToWord(generatedExam, 'De_On_Tap')}
                                            style={{ padding: '6px 12px', background: '#16a34a', color: 'white', border: 'none', borderRadius: '6px', fontSize: '12px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
                                        >
                                            <FileDown size={14} /> Tải file Word
                                        </button>
                                    )}
                                </div>
                                <div style={{ flex: 1, padding: '30px', overflowY: 'auto', whiteSpace: 'pre-wrap', fontFamily: 'Be Vietnam Pro', fontSize: '15px', lineHeight: '1.7', color: '#1e293b' }}>
                                    {generatedExam ? generatedExam : (
                                        <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: '#94a3b8' }}>
                                            <BrainCircuit size={48} style={{ opacity: 0.2, marginBottom: '20px' }} />
                                            <p>Vui lòng tải đề gốc và chạy phân tích để tạo nội dung.</p>
                                        </div>
                                    )}
                                </div>
                            </div>
                        </div>

                        </div>
                    )}
                    </>
                )}

             </div>
      </div>
      
      <style>{`
        .spin { animation: spin 1s linear infinite; }
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        table th, table td { vertical-align: middle; }
      `}</style>
    </div>
  );
};

const root = createRoot(document.getElementById('root')!);
root.render(<App />);