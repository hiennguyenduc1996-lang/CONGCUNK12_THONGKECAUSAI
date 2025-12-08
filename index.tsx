import React, { useState, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI } from "@google/genai";
import { Upload, FileText, Download, Loader2, Settings, Key, Eye, EyeOff, Calculator, FlaskConical, Languages, BrainCircuit, Table as TableIcon, X, User, School, BookOpen, ChevronRight, LayoutDashboard, FileSpreadsheet, RefreshCw, ArrowUpDown, ArrowUp, ArrowDown, FileDown, Filter, Palette, Monitor, Hourglass, Trophy, ClipboardList, Users, TrendingUp, Save, GraduationCap, FileOutput } from 'lucide-react';

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

// --- Ranking Types ---
interface RankingStudent {
    sbd: string;
    lastName: string;
    firstName: string;
    className: string;
}

interface PeriodScore {
    math: number | null;
    phys: number | null;
    chem: number | null;
    eng: number | null;
    bio: number | null;
}

// Map: Period (1-40) -> StudentSBD -> Scores
type ScoreDatabase = Record<number, Record<string, PeriodScore>>;

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

// Calculate Score for Part 2 (Group Questions)
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
  const p2 = config.parts.p2;
  
  if (index < p2.start || index > p2.end) return String(index);
  const relativeIndex = index - p2.start;
  const groupNum = Math.floor(relativeIndex / 4) + 1;
  const charCode = 97 + (relativeIndex % 4); 
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
                body { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.5; }
                p { margin-bottom: 10px; }
            </style>
        </head>
        <body>${content.replace(/\n/g, '<br>')}</body>
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
  const keysByVersion: Record<string, Record<number, string>> = {};

  for (let i = 1; i <= config.totalQuestions; i++) {
    questionStats[i] = 0;
    correctKeysForDisplay[i] = ''; 
  }

  data.forEach(row => {
    if (!row['StudentID'] && !row['LastName'] && !row['FirstName']) return;
    const version = String(row['Key Version'] || row['Exam Code'] || 'default').trim();
    let p1Score = 0; let p2Score = 0; let p3Score = 0;
    const details: Record<string, 'T' | 'F'> = {};
    const rawAnswers: Record<string, string> = {};

    const checkQuestion = (idx: number) => {
      const stCol = `Stu${idx}`;
      const keyCol = `PriKey${idx}`;
      const stAns = String(row[stCol] || '').trim().toUpperCase();
      let keyAns = String(row[keyCol] || '').trim().toUpperCase();

      if (!keyAns && keysByVersion[version] && keysByVersion[version][idx]) {
          keyAns = keysByVersion[version][idx];
      }
      if (keyAns) {
          if (!keysByVersion[version]) keysByVersion[version] = {};
          keysByVersion[version][idx] = keyAns;
          if (!correctKeysForDisplay[idx]) correctKeysForDisplay[idx] = keyAns;
      }
      rawAnswers[idx] = stAns;
      if (!keyAns) return false;
      return stAns === keyAns;
    };

    for (let i = config.parts.p1.start; i <= config.parts.p1.end; i++) {
      if (i === 0) continue;
      if (checkQuestion(i)) { p1Score += config.parts.p1.scorePerQ; details[i] = 'T'; } else { details[i] = 'F'; questionStats[i]++; }
    }
    if (config.parts.p2.end > 0) {
      for (let i = config.parts.p2.start; i <= config.parts.p2.end; i += 4) {
        let correctInGroup = 0;
        for (let j = 0; j < 4; j++) {
           const currentQ = i + j;
           if (currentQ > config.parts.p2.end) break;
           if (checkQuestion(currentQ)) { correctInGroup++; details[currentQ] = 'T'; } else { details[currentQ] = 'F'; questionStats[currentQ]++; }
        }
        p2Score += calculateGroupScore(correctInGroup);
      }
    }
    if (config.parts.p3.end > 0) {
      for (let i = config.parts.p3.start; i <= config.parts.p3.end; i++) {
         if (checkQuestion(i)) { p3Score += config.parts.p3.scorePerQ; details[i] = 'T'; } else { details[i] = 'F'; questionStats[i]++; }
      }
    }

    p1Score = Math.round(p1Score * 100) / 100;
    p2Score = Math.round(p2Score * 100) / 100;
    p3Score = Math.round(p3Score * 100) / 100;
    const totalScore = Math.round((p1Score + p2Score + p3Score) * 100) / 100;
    const fName = String(row['FirstName'] || '').trim();
    const lName = String(row['LastName'] || '').trim();

    results.push({
      sbd: String(row['StudentID'] || ''),
      firstName: fName,
      lastName: lName,
      name: `${fName} ${lName}`.trim(),
      code: version,
      rawAnswers,
      scores: { total: totalScore, p1: p1Score, p2: p2Score, p3: p3Score },
      details
    });
  });

  const stats: QuestionStat[] = [];
  const totalStudents = results.length;
  const isMultiVersion = Object.keys(keysByVersion).length > 1;

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


// --- Main App Component ---

const App = () => {
  const [activeSubject, setActiveSubject] = useState<'math' | 'science' | 'english' | 'it' | 'history' | 'ranking'>('math');
  const [activeTab, setActiveTab] = useState<'stats' | 'create'>('stats');
  
  // Data State
  const [data, setData] = useState<any[] | null>(null);
  const [processedResults, setProcessedResults] = useState<StudentResult[] | null>(null);
  const [stats, setStats] = useState<QuestionStat[] | null>(null);
  const [fileName, setFileName] = useState<string>("");

  // Ranking Tab State
  const [rankingTab, setRankingTab] = useState<'students' | 'scores' | 'summary' | 'report'>('students');
  const [rankingStudents, setRankingStudents] = useState<RankingStudent[]>(() => {
      const saved = localStorage.getItem('rankingStudents');
      return saved ? JSON.parse(saved) : [];
  });
  const [scoreDatabase, setScoreDatabase] = useState<ScoreDatabase>(() => {
      const saved = localStorage.getItem('scoreDatabase');
      return saved ? JSON.parse(saved) : {};
  });
  
  const [rankingFilter, setRankingFilter] = useState<'math'|'phys'|'chem'|'eng'|'bio'|'A'|'A1'|'B'|'All'>('All');
  const [rankingSort, setRankingSort] = useState<{key: string, direction: 'asc'|'desc'}>({key: 'score', direction: 'desc'});
  const [scoreInput, setScoreInput] = useState("");
  const [examPeriod, setExamPeriod] = useState<number>(1);
  const [selectedClass, setSelectedClass] = useState<string>("");
  const [selectedStudentForReport, setSelectedStudentForReport] = useState<RankingStudent | null>(null);

  // Stats Filter
  const [statsPartFilter, setStatsPartFilter] = useState<'all' | 'p1' | 'p2' | 'p3'>('all');
  const [questionCounts, setQuestionCounts] = useState<Record<number, number>>({});
  const [sortConfig, setSortConfig] = useState<{ key: string; direction: 'asc' | 'desc' } | null>(null);

  // Settings & Thresholds
  const [thresholds, setThresholds] = useState<ThresholdConfig>(() => {
    const saved = localStorage.getItem('thresholds');
    return saved ? JSON.parse(saved) : { lowCount: 5, highPercent: 40 };
  });
  const [customColors, setCustomColors] = useState<ColorConfig>(() => {
    const saved = localStorage.getItem('customColors');
    return saved ? JSON.parse(saved) : { lowError: DEFAULT_COLORS.yellow, highError: DEFAULT_COLORS.red };
  });

  // Exam Creation State
  const [examFile, setExamFile] = useState<DocFile | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedExam, setGeneratedExam] = useState<string>("");

  // Settings UI
  const [showSettings, setShowSettings] = useState(false);

  // Persistence
  useEffect(() => { localStorage.setItem('thresholds', JSON.stringify(thresholds)); }, [thresholds]);
  useEffect(() => { localStorage.setItem('customColors', JSON.stringify(customColors)); }, [customColors]);
  useEffect(() => { localStorage.setItem('rankingStudents', JSON.stringify(rankingStudents)); }, [rankingStudents]);
  useEffect(() => { localStorage.setItem('scoreDatabase', JSON.stringify(scoreDatabase)); }, [scoreDatabase]);

  // Handle File Upload (Excel/CSV) for ZipGrade
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
      if (activeSubject !== 'ranking') {
          const { results, stats } = processData(jsonData, activeSubject as any);
          setProcessedResults(results);
          setStats(stats);
      }
      setSortConfig(null);
    };
    reader.onload = (evt) => {
       const bstr = evt.target?.result;
       if (bstr) processBinary(bstr);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  // Re-process when subject changes (except ranking)
  useEffect(() => {
    if (activeSubject !== 'ranking' && data) {
      const { results, stats: newStats } = processData(data, activeSubject);
      setProcessedResults(results);
      setStats(newStats);
      setQuestionCounts({}); 
    }
  }, [activeSubject, data]);

  const handleSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') direction = 'desc';
    setSortConfig({ key, direction });
  };

  const sortedResults = useMemo(() => {
    if (!processedResults) return [];
    if (!sortConfig) return processedResults;
    const sorted = [...processedResults];
    sorted.sort((a, b) => {
      let aVal: any = ''; let bVal: any = '';
      if (sortConfig.key === 'name') { aVal = a.firstName; bVal = b.firstName; }
      else if (sortConfig.key === 'sbd') { aVal = a.sbd; bVal = b.sbd; }
      else if (sortConfig.key === 'total') { aVal = a.scores.total; bVal = b.scores.total; }
      else if (sortConfig.key === 'p1') { aVal = a.scores.p1; bVal = b.scores.p1; }
      else if (sortConfig.key === 'p2') { aVal = a.scores.p2; bVal = b.scores.p2; }
      else if (sortConfig.key === 'p3') { aVal = a.scores.p3; bVal = b.scores.p3; }
      if (aVal < bVal) return sortConfig.direction === 'asc' ? -1 : 1;
      if (aVal > bVal) return sortConfig.direction === 'asc' ? 1 : -1;
      return 0;
    });
    return sorted;
  }, [processedResults, sortConfig]);

  const summaryStats = useMemo(() => {
    if (!processedResults || processedResults.length === 0) return { min: 0, max: 0, avg: 0 };
    const scores = processedResults.map(r => r.scores.total);
    const sum = scores.reduce((a, b) => a + b, 0);
    return { min: Math.min(...scores), max: Math.max(...scores), avg: parseFloat((sum / scores.length).toFixed(2)) };
  }, [processedResults]);

  const p2Range = useMemo(() => {
    if (activeSubject === 'ranking') return { start: 0, end: 0 };
    return SUBJECTS_CONFIG[activeSubject].parts.p2;
  }, [activeSubject]);

  const filteredWrongStats = useMemo(() => {
      if (!stats) return [];
      const config = SUBJECTS_CONFIG[activeSubject === 'ranking' ? 'math' : activeSubject];
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

  const updateQuestionCount = (index: number, val: number) => setQuestionCounts(prev => ({ ...prev, [index]: val }));

  const handleExamFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    let content = "";
    if (file.type === "application/pdf") content = await fileToBase64(file);
    else if (file.name.endsWith(".docx") || file.name.endsWith(".doc")) content = await extractTextFromDocx(file);
    else {
      const reader = new FileReader();
      content = await new Promise((resolve) => { reader.onload = (e) => resolve(e.target?.result as string); reader.readAsText(file); });
    }
    setExamFile({ id: 'exam_orig', name: file.name, content, type: file.type === "application/pdf" ? 'pdf' : 'text' });
  };

  const generatePracticeExam = async () => {
    if (!examFile || !stats) return;
    setIsGenerating(true);
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      const requestDetails = filteredWrongStats.map(s => {
          const label = getPart2Label(s.index, activeSubject as any);
          const count = questionCounts[s.index] || 5; 
          return `- Dạng bài câu ${label}: tạo ${count} câu.`;
      }).join('\n');
      const prompt = `Bạn là một giáo viên chuyên nghiệp... (Prompt content truncated for brevity) ... ${requestDetails} ... Nội dung đề gốc: ${examFile.type === 'text' ? examFile.content : '(Xem PDF đính kèm)'}`;
      const parts: any[] = [{ text: prompt }];
      if (examFile.type === 'pdf') parts.push({ inlineData: { mimeType: 'application/pdf', data: examFile.content } });
      const response = await ai.models.generateContent({ model: 'gemini-2.5-flash', contents: { parts } });
      if (response.text) setGeneratedExam(response.text);
    } catch (e: any) { alert("Lỗi AI: " + e.message); } finally { setIsGenerating(false); }
  };

  const getCellColor = (wrongCount: number, wrongPercent: number) => {
    if (wrongCount === 0) return DEFAULT_COLORS.blue;
    if (wrongPercent > thresholds.highPercent) return customColors.highError;
    if (wrongCount < thresholds.lowCount) return customColors.lowError;
    return 'white';
  };
  const getScoreColor = (score: number) => { if (score >= 8) return '#16a34a'; if (score >= 5) return '#ca8a04'; return '#dc2626'; };
  const renderSortIcon = (key: string) => {
    if (!sortConfig || sortConfig.key !== key) return <ArrowUpDown size={12} style={{ opacity: 0.3 }} />;
    return sortConfig.direction === 'asc' ? <ArrowUp size={12} /> : <ArrowDown size={12} />;
  };

  // --- RANKING LOGIC ---

  const handleRankingStudentUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
          const bstr = evt.target?.result;
          if (bstr) {
              const wb = XLSX.read(bstr, { type: 'binary' });
              const ws = wb.Sheets[wb.SheetNames[0]];
              const data = XLSX.utils.sheet_to_json(ws, {header: 1});
              const students: RankingStudent[] = [];
              data.forEach((row: any, idx: number) => {
                  if (idx === 0) return; 
                  if (!row[0]) return;
                  students.push({ sbd: String(row[0]), lastName: String(row[1]||''), firstName: String(row[2]||''), className: String(row[3]||'') });
              });
              setRankingStudents(students);
          }
      };
      reader.readAsBinaryString(file);
      e.target.value = '';
  };

  const handleScoreFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
          const bstr = evt.target?.result;
          if (bstr) {
              const wb = XLSX.read(bstr, { type: 'binary' });
              let ws = wb.Sheets["DIEMKHOI"];
              if (!ws) {
                  alert("Không tìm thấy sheet tên 'DIEMKHOI' trong file! Vui lòng kiểm tra lại file Excel.");
                  return;
              }
              // Parse B4:J..
              // Row 4 is index 3. Columns B(1) to J(9).
              const data = XLSX.utils.sheet_to_json(ws, {header: 1, range: 3});
              
              const currentPeriodScores = { ...scoreDatabase[examPeriod] } || {};
              const newStudents: RankingStudent[] = [];
              const existingSBDs = new Set(rankingStudents.map(s => s.sbd));

              data.forEach((row: any) => {
                  const sbd = row[1] ? String(row[1]).trim() : "";
                  if (!sbd) return;

                  const parseScore = (val: any) => {
                      if (typeof val === 'number') return val;
                      if (!val) return null;
                      const float = parseFloat(String(val).replace(',', '.'));
                      return isNaN(float) ? null : float;
                  };

                  // Update Scores
                  currentPeriodScores[sbd] = {
                      math: parseScore(row[5]), // F
                      phys: parseScore(row[6]), // G
                      chem: parseScore(row[7]), // H
                      eng: parseScore(row[8]),  // I
                      bio: parseScore(row[9]),  // J
                  };

                  // Auto-add student if not exists
                  if (!existingSBDs.has(sbd)) {
                      newStudents.push({
                          sbd: sbd,
                          lastName: row[2] ? String(row[2]).trim() : "",
                          firstName: row[3] ? String(row[3]).trim() : "",
                          className: row[4] ? String(row[4]).trim() : "",
                      });
                      existingSBDs.add(sbd);
                  }
              });

              setScoreDatabase({ ...scoreDatabase, [examPeriod]: currentPeriodScores });
              if (newStudents.length > 0) {
                  setRankingStudents(prev => [...prev, ...newStudents]);
              }
              alert(`Đã cập nhật điểm từ sheet 'DIEMKHOI' cho lần kiểm tra ${examPeriod}. Thêm mới ${newStudents.length} học sinh.`);
          }
      };
      reader.readAsBinaryString(file);
      e.target.value = '';
  };

  const processScoreInput = () => {
      if (!scoreInput.trim()) return;
      const rows = scoreInput.trim().split('\n');
      const currentPeriodScores = { ...scoreDatabase[examPeriod] } || {};
      const newStudents: RankingStudent[] = [];
      const existingSBDs = new Set(rankingStudents.map(s => s.sbd));
      
      rows.forEach(row => {
          const cols = row.split('\t');
          if (cols.length < 1) return;
          const sbd = String(cols[0]).trim();
          if (!sbd) return;
          
          currentPeriodScores[sbd] = {
              math: cols[4] ? parseFloat(cols[4].replace(',', '.')) : null,
              phys: cols[5] ? parseFloat(cols[5].replace(',', '.')) : null,
              chem: cols[6] ? parseFloat(cols[6].replace(',', '.')) : null,
              eng: cols[7] ? parseFloat(cols[7].replace(',', '.')) : null,
              bio: cols[8] ? parseFloat(cols[8].replace(',', '.')) : null,
          };

          if (!existingSBDs.has(sbd)) {
              newStudents.push({
                  sbd: sbd,
                  lastName: cols[1] ? cols[1].trim() : "",
                  firstName: cols[2] ? cols[2].trim() : "",
                  className: cols[3] ? cols[3].trim() : "",
              });
              existingSBDs.add(sbd);
          }
      });

      setScoreDatabase({ ...scoreDatabase, [examPeriod]: currentPeriodScores });
      if (newStudents.length > 0) {
          setRankingStudents(prev => [...prev, ...newStudents]);
      }
      setScoreInput("");
      alert(`Đã nhập điểm cho lần kiểm tra ${examPeriod}! Thêm mới ${newStudents.length} học sinh.`);
  };

  const updateIndividualScore = (sbd: string, subject: keyof PeriodScore, value: string) => {
      const val = value === '' ? null : parseFloat(value);
      const currentPeriodScores = { ...(scoreDatabase[examPeriod] || {}) };
      const studentScores = { ...(currentPeriodScores[sbd] || { math: null, phys: null, chem: null, eng: null, bio: null }) };
      
      (studentScores as any)[subject] = val;
      currentPeriodScores[sbd] = studentScores;
      
      setScoreDatabase({ ...scoreDatabase, [examPeriod]: currentPeriodScores });
  };

  const saveScoreDatabase = () => {
      localStorage.setItem('scoreDatabase', JSON.stringify(scoreDatabase));
      alert("Đã lưu dữ liệu điểm thành công!");
  };

  const getFilteredRanking = useMemo(() => {
      let data = rankingStudents.map(st => {
          // Calculate Average for each subject across all available periods
          const subjectTotals = { math: 0, phys: 0, chem: 0, eng: 0, bio: 0 };
          const subjectCounts = { math: 0, phys: 0, chem: 0, eng: 0, bio: 0 };

          Object.values(scoreDatabase).forEach(periodScores => {
              const s = periodScores[st.sbd];
              if(s) {
                  if(s.math !== null) { subjectTotals.math += s.math; subjectCounts.math++; }
                  if(s.phys !== null) { subjectTotals.phys += s.phys; subjectCounts.phys++; }
                  if(s.chem !== null) { subjectTotals.chem += s.chem; subjectCounts.chem++; }
                  if(s.eng !== null) { subjectTotals.eng += s.eng; subjectCounts.eng++; }
                  if(s.bio !== null) { subjectTotals.bio += s.bio; subjectCounts.bio++; }
              }
          });

          // Compute averages
          const scores = {
              math: subjectCounts.math ? parseFloat((subjectTotals.math / subjectCounts.math).toFixed(2)) : null,
              phys: subjectCounts.phys ? parseFloat((subjectTotals.phys / subjectCounts.phys).toFixed(2)) : null,
              chem: subjectCounts.chem ? parseFloat((subjectTotals.chem / subjectCounts.chem).toFixed(2)) : null,
              eng: subjectCounts.eng ? parseFloat((subjectTotals.eng / subjectCounts.eng).toFixed(2)) : null,
              bio: subjectCounts.bio ? parseFloat((subjectTotals.bio / subjectCounts.bio).toFixed(2)) : null,
          };

          const m = scores.math || 0;
          const p = scores.phys || 0;
          const c = scores.chem || 0;
          const e = scores.eng || 0;
          const b = scores.bio || 0;

          // Block scores are Sum of Subject Averages
          const blockA = (scores.math===null || scores.phys===null || scores.chem===null) ? 0 : m + p + c;
          const blockA1 = (scores.math===null || scores.phys===null || scores.eng===null) ? 0 : m + p + e;
          const blockB = (scores.math===null || scores.chem===null || scores.bio===null) ? 0 : m + c + b;
          const maxBlock = Math.max(blockA, blockA1, blockB);

          let sortValue = 0;
          if (rankingFilter === 'math') sortValue = m;
          else if (rankingFilter === 'phys') sortValue = p;
          else if (rankingFilter === 'chem') sortValue = c;
          else if (rankingFilter === 'eng') sortValue = e;
          else if (rankingFilter === 'bio') sortValue = b;
          else if (rankingFilter === 'A') sortValue = blockA;
          else if (rankingFilter === 'A1') sortValue = blockA1;
          else if (rankingFilter === 'B') sortValue = blockB;
          else if (rankingFilter === 'All') sortValue = maxBlock;

          return { ...st, scores, blockA, blockA1, blockB, maxBlock, sortValue };
      });

      data.sort((a, b) => {
          if (rankingSort.key === 'sbd') return rankingSort.direction === 'asc' ? a.sbd.localeCompare(b.sbd) : b.sbd.localeCompare(a.sbd);
          if (rankingSort.key === 'name') {
              const nameA = a.firstName.toLowerCase(); const nameB = b.firstName.toLowerCase();
              if (nameA < nameB) return rankingSort.direction === 'asc' ? -1 : 1;
              if (nameA > nameB) return rankingSort.direction === 'asc' ? 1 : -1;
              return 0;
          }
          if (rankingSort.key === 'score') return rankingSort.direction === 'asc' ? a.sortValue - b.sortValue : b.sortValue - a.sortValue;
          return 0;
      });
      return data;
  }, [rankingStudents, scoreDatabase, rankingFilter, rankingSort]);

  const classStats = useMemo(() => {
      const stats: Record<string, number> = {};
      rankingStudents.forEach(st => {
          stats[st.className] = (stats[st.className] || 0) + 1;
      });
      return stats;
  }, [rankingStudents]);

  // --- REPORT CARD LOGIC ---
  const reportData = useMemo(() => {
      if (!selectedStudentForReport) return null;
      const st = selectedStudentForReport;
      const className = st.className.toUpperCase();
      
      let subjects: {key: keyof PeriodScore, label: string}[] = [];
      let type: 'A' | 'B' | 'A1' | 'Unknown' = 'Unknown';

      if (className.startsWith('12A')) { type = 'A'; subjects = [{key:'math', label:'Toán'}, {key:'phys', label:'Lí'}, {key:'chem', label:'Hóa'}]; }
      else if (className.startsWith('12B')) { type = 'B'; subjects = [{key:'math', label:'Toán'}, {key:'chem', label:'Hóa'}, {key:'bio', label:'Sinh'}]; }
      else if (className.startsWith('12E')) { type = 'A1'; subjects = [{key:'math', label:'Toán'}, {key:'phys', label:'Lí'}, {key:'eng', label:'Anh'}]; }
      else {
          // Default fallback
          subjects = [{key:'math', label:'Toán'}, {key:'phys', label:'Lí'}, {key:'chem', label:'Hóa'}, {key:'eng', label:'Anh'}, {key:'bio', label:'Sinh'}];
      }

      const rows = [];
      for (let i = 1; i <= 40; i++) {
          const scores = scoreDatabase[i]?.[st.sbd] || {};
          const rowData: any = { period: i };
          let rowSum = 0;
          let count = 0;
          subjects.forEach(sub => {
              const val = (scores as any)[sub.key];
              rowData[sub.key] = val;
              if (val !== null && val !== undefined) { rowSum += val; count++; }
          });
          rowData.avg = count === subjects.length ? parseFloat((rowSum).toFixed(2)) : null; 
          rowData.blockSum = count === 3 ? parseFloat(rowSum.toFixed(2)) : null; 
          rows.push(rowData);
      }

      // Vertical Averages
      const colAvgs: any = {};
      subjects.forEach(sub => {
          let sum = 0; let c = 0;
          rows.forEach(r => { if (r[sub.key] !== null && r[sub.key] !== undefined) { sum += r[sub.key]; c++; } });
          colAvgs[sub.key] = c ? parseFloat((sum/c).toFixed(2)) : null;
      });
      // Avg of Block Sums
      let bSum = 0; let bC = 0;
      rows.forEach(r => { if (r.blockSum !== null) { bSum += r.blockSum; bC++; } });
      colAvgs.blockSum = bC ? parseFloat((bSum/bC).toFixed(2)) : null;

      return { student: st, type, subjects, rows, colAvgs };
  }, [selectedStudentForReport, scoreDatabase]);

  const exportReportCard = () => {
      if (!reportData) return;
      const { student, subjects, rows, colAvgs } = reportData;
      
      const wsData = [
          [`Trường THCS và THPT Nguyễn Khuyến Bình Dương`],
          [`ĐIỂM TRUNG BÌNH HỌC SINH: ${student.firstName} ${student.lastName} - Lớp: ${student.className}`],
          [],
          ['Lần KT', ...subjects.map(s => s.label), 'Tổng Khối']
      ];

      rows.forEach(r => {
          const row = [r.period];
          subjects.forEach(s => row.push(r[s.key] ?? ''));
          row.push(r.blockSum ?? '');
          wsData.push(row);
      });

      const avgRow = ['TB Môn'];
      subjects.forEach(s => avgRow.push(colAvgs[s.key] ?? ''));
      avgRow.push(colAvgs.blockSum ?? '');
      wsData.push(avgRow);

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      
      // Merge headers
      ws['!merges'] = [
          { s: { r: 0, c: 0 }, e: { r: 0, c: subjects.length + 1 } },
          { s: { r: 1, c: 0 }, e: { r: 1, c: subjects.length + 1 } }
      ];

      XLSX.utils.book_append_sheet(wb, ws, "PhieuDiem");
      XLSX.writeFile(wb, `PhieuDiem_${student.sbd}.xlsx`);
  };

  // --- UI RENDER ---

  const renderRankingContent = () => {
      return (
          <div style={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
              {/* Sidebar */}
              <div style={{ width: '260px', background: 'white', borderRight: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column' }}>
                  <div style={{ padding: '20px', fontWeight: 700, color: '#334155', borderBottom: '1px solid #f1f5f9' }}>CHỨC NĂNG</div>
                  <div style={{ padding: '10px', display: 'flex', flexDirection: 'column', gap: '5px' }}>
                      <button onClick={() => setRankingTab('students')} style={{ padding: '12px 15px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left', background: rankingTab === 'students' ? '#eff6ff' : 'transparent', color: rankingTab === 'students' ? '#1d4ed8' : '#64748b', fontWeight: rankingTab === 'students' ? 600 : 500, display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <Users size={18} /> Danh sách học sinh
                      </button>
                      <button onClick={() => setRankingTab('scores')} style={{ padding: '12px 15px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left', background: rankingTab === 'scores' ? '#eff6ff' : 'transparent', color: rankingTab === 'scores' ? '#1d4ed8' : '#64748b', fontWeight: rankingTab === 'scores' ? 600 : 500, display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <ClipboardList size={18} /> Dữ liệu điểm
                      </button>
                      <button onClick={() => setRankingTab('summary')} style={{ padding: '12px 15px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left', background: rankingTab === 'summary' ? '#eff6ff' : 'transparent', color: rankingTab === 'summary' ? '#1d4ed8' : '#64748b', fontWeight: rankingTab === 'summary' ? 600 : 500, display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <TrendingUp size={18} /> Tổng kết
                      </button>
                      <button onClick={() => setRankingTab('report')} style={{ padding: '12px 15px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left', background: rankingTab === 'report' ? '#eff6ff' : 'transparent', color: rankingTab === 'report' ? '#1d4ed8' : '#64748b', fontWeight: rankingTab === 'report' ? 600 : 500, display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <FileOutput size={18} /> Phiếu điểm học sinh
                      </button>
                  </div>
              </div>

              {/* Content */}
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: '#f8fafc' }}>
                  {rankingTab === 'students' && (
                      <div style={{ flex: 1, padding: '20px', overflow: 'auto' }}>
                          <div style={{ maxWidth: '900px', margin: '0 auto', background: 'white', borderRadius: '12px', padding: '30px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', height: '100%' }}>
                              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                                  <div>
                                      <h3 style={{ margin: 0, color: '#1e3a8a' }}>Danh sách học sinh</h3>
                                      <p style={{ margin: '5px 0 0 0', fontSize: '13px', color: '#64748b' }}>Tổng số: <strong>{rankingStudents.length}</strong> học sinh</p>
                                  </div>
                                  <div style={{ cursor: 'pointer', padding: '10px 20px', background: '#eff6ff', borderRadius: '8px', border: '1px solid #bfdbfe', display: 'flex', gap: '8px', alignItems: 'center', color: '#1d4ed8', fontWeight: 500 }} onClick={() => document.getElementById('rank-student-upload')?.click()}>
                                      <Upload size={18} /> Tải file Excel
                                      <input type="file" id="rank-student-upload" accept=".xlsx" style={{display:'none'}} onChange={handleRankingStudentUpload} />
                                  </div>
                              </div>

                              {/* Class Stats */}
                              {rankingStudents.length > 0 && (
                                  <div style={{ marginBottom: '20px', padding: '15px', background: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0', display: 'flex', flexWrap: 'wrap', gap: '15px' }}>
                                      {Object.entries(classStats).sort().map(([cls, count]) => (
                                          <div key={cls} style={{ background: 'white', padding: '6px 12px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '12px', color: '#334155' }}>
                                              Lớp <strong>{cls}</strong>: {count}
                                          </div>
                                      ))}
                                  </div>
                              )}

                              <div style={{ flex: 1, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '8px' }}>
                                  <table style={{ width: '100%', fontSize: '13px', borderCollapse: 'collapse' }}>
                                      <thead style={{ position: 'sticky', top: 0, background: '#f1f5f9', zIndex: 1 }}>
                                          <tr>
                                              <th style={{ width: '50px', padding: '10px' }}>STT</th>
                                              <th style={{ textAlign: 'left', padding: '10px' }}>SBD</th>
                                              <th style={{ textAlign: 'left', padding: '10px' }}>Họ</th>
                                              <th style={{ textAlign: 'left', padding: '10px' }}>Tên</th>
                                              <th style={{ textAlign: 'center', padding: '10px' }}>Lớp</th>
                                          </tr>
                                      </thead>
                                      <tbody>
                                          {rankingStudents.map((st, idx) => (
                                              <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                                                  <td style={{ textAlign: 'center', padding: '8px' }}>{idx + 1}</td>
                                                  <td style={{ padding: '8px' }}>{st.sbd}</td>
                                                  <td style={{ padding: '8px' }}>{st.lastName}</td>
                                                  <td style={{ padding: '8px' }}>{st.firstName}</td>
                                                  <td style={{ textAlign: 'center', padding: '8px' }}>{st.className}</td>
                                              </tr>
                                          ))}
                                          {rankingStudents.length === 0 && <tr><td colSpan={5} style={{ textAlign: 'center', padding: '30px', color: '#94a3b8' }}>Chưa có dữ liệu</td></tr>}
                                      </tbody>
                                  </table>
                              </div>
                          </div>
                      </div>
                  )}

                  {rankingTab === 'scores' && (
                      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                          {/* 40 Tabs Scrollable */}
                          <div style={{ padding: '10px 20px', background: 'white', borderBottom: '1px solid #e2e8f0', display: 'flex', gap: '6px', overflowX: 'auto', flexShrink: 0 }}>
                              {Array.from({ length: 40 }, (_, i) => i + 1).map(num => (
                                  <button
                                      key={num}
                                      onClick={() => setExamPeriod(num)}
                                      style={{
                                          minWidth: '60px', padding: '8px', borderRadius: '6px', border: '1px solid', cursor: 'pointer', fontSize: '13px',
                                          background: examPeriod === num ? '#1e3a8a' : 'white',
                                          color: examPeriod === num ? 'white' : '#64748b',
                                          borderColor: examPeriod === num ? '#1e3a8a' : '#e2e8f0',
                                          fontWeight: examPeriod === num ? 600 : 400
                                      }}
                                  >
                                      Lần {num}
                                  </button>
                              ))}
                          </div>

                          <div style={{ flex: 1, padding: '20px', overflow: 'auto', display: 'flex', flexDirection: 'column' }}>
                              <div style={{ background: 'white', borderRadius: '12px', padding: '20px', border: '1px solid #e2e8f0', flex: 1, display: 'flex', flexDirection: 'column' }}>
                                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '15px' }}>
                                      <h3 style={{ margin: 0, color: '#1e3a8a' }}>Nhập điểm - Lần kiểm tra {examPeriod}</h3>
                                      <div style={{ display: 'flex', gap: '10px' }}>
                                          <div style={{ cursor: 'pointer', padding: '8px 16px', background: '#eff6ff', borderRadius: '6px', border: '1px solid #bfdbfe', display: 'flex', gap: '8px', alignItems: 'center', color: '#1d4ed8', fontWeight: 600 }} onClick={() => document.getElementById('score-file-upload')?.click()}>
                                              <Upload size={16} /> Tải Excel (Sheet DIEMKHOI)
                                              <input type="file" id="score-file-upload" accept=".xlsx" style={{display:'none'}} onChange={handleScoreFileUpload} />
                                          </div>
                                          <button onClick={saveScoreDatabase} style={{ padding: '8px 16px', background: '#16a34a', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 600 }}>
                                              <Save size={16} /> Lưu lại
                                          </button>
                                      </div>
                                  </div>

                                  <div style={{ marginBottom: '20px' }}>
                                      <textarea
                                          style={{ width: '100%', height: '80px', padding: '10px', borderRadius: '8px', border: '1px solid #cbd5e1', fontFamily: 'monospace', fontSize: '12px', resize: 'none' }}
                                          placeholder="Dán dữ liệu Excel vào đây (9 cột: SBD | Họ | Tên | Lớp | Toán | Lí | Hóa | Anh | Sinh)"
                                          value={scoreInput}
                                          onChange={(e) => setScoreInput(e.target.value)}
                                          onKeyDown={(e) => { if(e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); processScoreInput(); } }}
                                      />
                                      <div style={{ fontSize: '11px', color: '#64748b', marginTop: '4px' }}>Nhấn Enter để xử lý dữ liệu dán</div>
                                  </div>

                                  <div style={{ flex: 1, overflow: 'auto', border: '1px solid #e2e8f0', borderRadius: '8px' }}>
                                      <table style={{ width: '100%', fontSize: '13px', borderCollapse: 'collapse' }}>
                                          <thead style={{ position: 'sticky', top: 0, background: '#f1f5f9', zIndex: 1 }}>
                                              <tr>
                                                  <th style={{ width: '40px', padding: '8px' }}>STT</th>
                                                  <th style={{ width: '80px', padding: '8px' }}>SBD</th>
                                                  <th style={{ textAlign: 'left', padding: '8px' }}>Họ Tên</th>
                                                  <th style={{ width: '60px', padding: '8px' }}>Lớp</th>
                                                  <th style={{ width: '60px', padding: '8px' }}>Toán</th>
                                                  <th style={{ width: '60px', padding: '8px' }}>Lí</th>
                                                  <th style={{ width: '60px', padding: '8px' }}>Hóa</th>
                                                  <th style={{ width: '60px', padding: '8px' }}>Anh</th>
                                                  <th style={{ width: '60px', padding: '8px' }}>Sinh</th>
                                              </tr>
                                          </thead>
                                          <tbody>
                                              {rankingStudents.length > 0 ? rankingStudents.map((st, idx) => {
                                                  const scores = scoreDatabase[examPeriod]?.[st.sbd] || {};
                                                  return (
                                                      <tr key={st.sbd} style={{ borderBottom: '1px solid #f1f5f9' }}>
                                                          <td style={{ textAlign: 'center' }}>{idx + 1}</td>
                                                          <td style={{ textAlign: 'center' }}>{st.sbd}</td>
                                                          <td style={{ padding: '8px' }}>{st.lastName} {st.firstName}</td>
                                                          <td style={{ textAlign: 'center' }}>{st.className}</td>
                                                          {['math', 'phys', 'chem', 'eng', 'bio'].map(sub => (
                                                              <td key={sub} style={{ padding: 0 }}>
                                                                  <input
                                                                      type="number"
                                                                      step="0.1"
                                                                      style={{ width: '100%', height: '100%', border: 'none', textAlign: 'center', padding: '8px', outline: 'none', background: 'transparent' }}
                                                                      value={(scores as any)[sub] ?? ''}
                                                                      onChange={(e) => updateIndividualScore(st.sbd, sub as any, e.target.value)}
                                                                  />
                                                              </td>
                                                          ))}
                                                      </tr>
                                                  );
                                              }) : (
                                                  <tr>
                                                      <td colSpan={9} style={{ textAlign: 'center', padding: '30px', color: '#94a3b8' }}>
                                                          Chưa có học sinh. Vui lòng tải "Danh sách học sinh" hoặc nhập điểm để tự động thêm.
                                                      </td>
                                                  </tr>
                                              )}
                                          </tbody>
                                      </table>
                                  </div>
                              </div>
                          </div>
                      </div>
                  )}

                  {rankingTab === 'summary' && (
                      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                          <div style={{ padding: '15px 20px', background: 'white', borderBottom: '1px solid #e2e8f0', display: 'flex', gap: '8px', overflowX: 'auto', alignItems: 'center' }}>
                              {['math', 'phys', 'chem', 'eng', 'bio'].map(k => (
                                  <button key={k} onClick={() => setRankingFilter(k as any)} style={{ padding: '6px 14px', borderRadius: '99px', border: '1px solid', fontSize: '13px', cursor: 'pointer', borderColor: rankingFilter === k ? '#2563eb' : '#e2e8f0', background: rankingFilter === k ? '#eff6ff' : 'white', color: rankingFilter === k ? '#1d4ed8' : '#64748b', fontWeight: 600 }}>
                                      {k === 'math' ? 'Toán' : k === 'phys' ? 'Lí' : k === 'chem' ? 'Hóa' : k === 'eng' ? 'Anh' : 'Sinh'}
                                  </button>
                              ))}
                              <div style={{ width: '1px', height: '20px', background: '#e2e8f0', margin: '0 5px' }}></div>
                              {['A1', 'A', 'B', 'All'].map(k => (
                                  <button key={k} onClick={() => setRankingFilter(k as any)} style={{ padding: '6px 14px', borderRadius: '99px', border: '1px solid', fontSize: '13px', cursor: 'pointer', borderColor: rankingFilter === k ? '#7c3aed' : '#e2e8f0', background: rankingFilter === k ? '#f5f3ff' : 'white', color: rankingFilter === k ? '#7c3aed' : '#64748b', fontWeight: 600 }}>
                                      {k === 'All' ? 'Toàn Khối' : `Khối ${k}`}
                                  </button>
                              ))}
                          </div>
                          <div style={{ flex: 1, padding: '20px', overflow: 'auto' }}>
                              <div style={{ background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                                  <table style={{ width: '100%', fontSize: '13px', borderCollapse: 'collapse' }}>
                                      <thead style={{ background: '#f1f5f9' }}>
                                          <tr>
                                              <th style={{ width: '50px', padding: '10px' }}>#</th>
                                              <th style={{ textAlign: 'left', padding: '10px', cursor:'pointer' }} onClick={() => setRankingSort({key:'sbd', direction: rankingSort.direction==='asc'?'desc':'asc'})}>SBD {renderSortIcon('sbd')}</th>
                                              <th style={{ textAlign: 'left', padding: '10px', cursor:'pointer' }} onClick={() => setRankingSort({key:'name', direction: rankingSort.direction==='asc'?'desc':'asc'})}>Họ và Tên {renderSortIcon('name')}</th>
                                              <th style={{ width: '60px', padding: '10px' }}>Lớp</th>
                                              
                                              {/* Dynamic Headers */}
                                              {['math','phys','chem','eng','bio'].includes(rankingFilter) && <th>Điểm TB</th>}
                                              {(rankingFilter === 'A' || rankingFilter === 'All') && <th style={{background:'#eff6ff', color:'#1d4ed8'}}>Khối A</th>}
                                              {(rankingFilter === 'A1' || rankingFilter === 'All') && <th style={{background:'#eff6ff', color:'#1d4ed8'}}>Khối A1</th>}
                                              {(rankingFilter === 'B' || rankingFilter === 'All') && <th style={{background:'#eff6ff', color:'#1d4ed8'}}>Khối B</th>}
                                              {rankingFilter === 'All' && <th style={{background:'#fef3c7', color:'#b45309'}}>Cao nhất</th>}
                                          </tr>
                                      </thead>
                                      <tbody>
                                          {getFilteredRanking.map((st, idx) => (
                                              <tr key={st.sbd} style={{ borderBottom: '1px solid #f1f5f9' }}>
                                                  <td style={{ textAlign: 'center', padding: '8px' }}>{idx + 1}</td>
                                                  <td style={{ padding: '8px' }}>{st.sbd}</td>
                                                  <td style={{ padding: '8px' }}>{st.lastName} {st.firstName}</td>
                                                  <td style={{ textAlign: 'center' }}>{st.className}</td>
                                                  
                                                  {['math','phys','chem','eng','bio'].includes(rankingFilter) && <td style={{textAlign:'center'}}>{st.sortValue}</td>}

                                                  {(rankingFilter === 'A' || rankingFilter === 'All') && <td style={{textAlign:'center', fontWeight:700, color:'#1d4ed8', background:'#eff6ff'}}>{st.blockA > 0 ? st.blockA.toFixed(2) : '-'}</td>}
                                                  {(rankingFilter === 'A1' || rankingFilter === 'All') && <td style={{textAlign:'center', fontWeight:700, color:'#1d4ed8', background:'#eff6ff'}}>{st.blockA1 > 0 ? st.blockA1.toFixed(2) : '-'}</td>}
                                                  {(rankingFilter === 'B' || rankingFilter === 'All') && <td style={{textAlign:'center', fontWeight:700, color:'#1d4ed8', background:'#eff6ff'}}>{st.blockB > 0 ? st.blockB.toFixed(2) : '-'}</td>}
                                                  {rankingFilter === 'All' && <td style={{textAlign:'center', fontWeight:700, color:'#b45309', background:'#fef3c7'}}>{st.maxBlock > 0 ? st.maxBlock.toFixed(2) : '-'}</td>}
                                              </tr>
                                          ))}
                                          {getFilteredRanking.length === 0 && <tr><td colSpan={10} style={{textAlign:'center', padding:'20px', color:'#94a3b8'}}>Chưa có dữ liệu tính toán</td></tr>}
                                      </tbody>
                                  </table>
                              </div>
                          </div>
                      </div>
                  )}

                  {rankingTab === 'report' && (
                      <div style={{ flex: 1, padding: '20px', overflow: 'auto', display: 'flex', gap: '20px' }}>
                          <div style={{ width: '300px', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', padding: '20px', display: 'flex', flexDirection: 'column' }}>
                              <h4 style={{ margin: '0 0 15px 0', color: '#1e3a8a' }}>Chọn Học Sinh</h4>
                              <div style={{ marginBottom: '15px' }}>
                                  <label style={{ fontSize: '13px', fontWeight: 600, color: '#64748b' }}>Lớp</label>
                                  <select 
                                      style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '5px' }}
                                      value={selectedClass}
                                      onChange={(e) => setSelectedClass(e.target.value)}
                                  >
                                      <option value="">-- Chọn lớp --</option>
                                      {Object.keys(classStats).sort().map(c => <option key={c} value={c}>{c}</option>)}
                                  </select>
                              </div>
                              <div style={{ flex: 1, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '6px' }}>
                                  {rankingStudents.filter(s => selectedClass ? s.className === selectedClass : true).map(s => (
                                      <div 
                                          key={s.sbd} 
                                          onClick={() => setSelectedStudentForReport(s)}
                                          style={{ padding: '8px 12px', borderBottom: '1px solid #f1f5f9', cursor: 'pointer', background: selectedStudentForReport?.sbd === s.sbd ? '#eff6ff' : 'white', fontSize: '13px' }}
                                      >
                                          {s.firstName} {s.lastName}
                                      </div>
                                  ))}
                              </div>
                          </div>

                          <div style={{ flex: 1, background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', padding: '30px', overflow: 'auto' }}>
                              {reportData ? (
                                  <div>
                                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '30px' }}>
                                          <div>
                                              <div style={{ fontSize: '14px', fontWeight: 600, textTransform: 'uppercase', color: '#64748b' }}>Trường THCS và THPT Nguyễn Khuyến Bình Dương</div>
                                              <h2 style={{ margin: '10px 0', color: '#1e3a8a' }}>ĐIỂM TRUNG BÌNH HỌC SINH</h2>
                                              <div style={{ fontSize: '16px', fontWeight: 600 }}>{reportData.student.firstName} {reportData.student.lastName}</div>
                                              <div style={{ fontSize: '14px' }}>Lớp: {reportData.student.className} (Khối {reportData.type})</div>
                                          </div>
                                          <button onClick={exportReportCard} style={{ padding: '10px 20px', background: '#16a34a', color: 'white', border: 'none', borderRadius: '8px', cursor: 'pointer', display: 'flex', gap: '8px', alignItems: 'center', fontWeight: 600 }}>
                                              <Download size={18} /> Xuất Excel
                                          </button>
                                      </div>

                                      <table style={{ width: '100%', fontSize: '13px', borderCollapse: 'collapse', textAlign: 'center' }}>
                                          <thead>
                                              <tr>
                                                  <th style={{ padding: '10px', background: '#f1f5f9', border: '1px solid #cbd5e1' }}>Lần KT</th>
                                                  {reportData.subjects.map(s => <th key={s.key} style={{ padding: '10px', background: '#f1f5f9', border: '1px solid #cbd5e1' }}>{s.label}</th>)}
                                                  <th style={{ padding: '10px', background: '#e0f2fe', border: '1px solid #cbd5e1', color: '#0369a1' }}>Tổng Khối</th>
                                              </tr>
                                          </thead>
                                          <tbody>
                                              {reportData.rows.map(r => (
                                                  <tr key={r.period}>
                                                      <td style={{ padding: '8px', border: '1px solid #cbd5e1' }}>Lần {r.period}</td>
                                                      {reportData.subjects.map(s => <td key={s.key} style={{ padding: '8px', border: '1px solid #cbd5e1' }}>{r[s.key]}</td>)}
                                                      <td style={{ padding: '8px', border: '1px solid #cbd5e1', fontWeight: 700, color: '#0369a1', background: '#f0f9ff' }}>{r.blockSum}</td>
                                                  </tr>
                                              ))}
                                              <tr style={{ background: '#f8fafc', fontWeight: 700 }}>
                                                  <td style={{ padding: '10px', border: '1px solid #cbd5e1' }}>TB Môn</td>
                                                  {reportData.subjects.map(s => <td key={s.key} style={{ padding: '10px', border: '1px solid #cbd5e1' }}>{reportData.colAvgs[s.key]}</td>)}
                                                  <td style={{ padding: '10px', border: '1px solid #cbd5e1', color: '#0369a1' }}>{reportData.colAvgs.blockSum}</td>
                                              </tr>
                                          </tbody>
                                      </table>
                                  </div>
                              ) : (
                                  <div style={{ height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#94a3b8', flexDirection: 'column' }}>
                                      <GraduationCap size={48} style={{ opacity: 0.2, marginBottom: '20px' }} />
                                      Chọn học sinh để xem phiếu điểm
                                  </div>
                              )}
                          </div>
                      </div>
                  )}
              </div>
          </div>
      );
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', fontFamily: 'Inter, sans-serif', background: '#f8fafc', color: '#1e293b' }}>
      <div style={{ height: '60px', background: 'white', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', padding: '0 20px', justifyContent: 'space-between', flexShrink: 0 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ width: '32px', height: '32px', background: 'linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%)', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'white' }}>
                    <GraduationCap size={20} />
                </div>
                <h1 style={{ fontSize: '18px', fontWeight: 700, color: '#1e3a8a', margin: 0 }}>ZipGrade Analytics AI</h1>
              </div>
              
              <div style={{ width: '1px', height: '24px', background: '#e2e8f0' }}></div>
              
              <div style={{ display: 'flex', gap: '4px', background: '#f1f5f9', padding: '4px', borderRadius: '8px' }}>
                {Object.values(SUBJECTS_CONFIG).map(sub => (
                    <button key={sub.id} onClick={() => setActiveSubject(sub.id as any)} style={{ padding: '6px 12px', borderRadius: '6px', border: 'none', fontSize: '13px', fontWeight: 600, cursor: 'pointer', background: activeSubject === sub.id ? 'white' : 'transparent', color: activeSubject === sub.id ? '#1e3a8a' : '#64748b', boxShadow: activeSubject === sub.id ? '0 1px 2px rgba(0,0,0,0.1)' : 'none', transition: 'all 0.2s' }}>
                        {sub.name}
                    </button>
                ))}
                <button onClick={() => setActiveSubject('ranking')} style={{ padding: '6px 12px', borderRadius: '6px', border: 'none', fontSize: '13px', fontWeight: 600, cursor: 'pointer', background: activeSubject === 'ranking' ? 'white' : 'transparent', color: activeSubject === 'ranking' ? '#1e3a8a' : '#64748b', boxShadow: activeSubject === 'ranking' ? '0 1px 2px rgba(0,0,0,0.1)' : 'none', transition: 'all 0.2s', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <Trophy size={14} /> Xếp hạng
                </button>
              </div>
          </div>
          
          <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
              <button onClick={() => setShowSettings(true)} style={{ padding: '8px', borderRadius: '8px', border: '1px solid #e2e8f0', background: 'white', color: '#64748b', cursor: 'pointer' }}>
                  <Settings size={20} />
              </button>
          </div>
      </div>

      {activeSubject === 'ranking' && renderRankingContent()}

      {activeSubject !== 'ranking' && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ padding: '15px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ display: 'flex', gap: '8px' }}>
                      <button onClick={() => setActiveTab('stats')} style={{ padding: '10px 20px', borderRadius: '99px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 600, background: activeTab === 'stats' ? '#1e3a8a' : '#e2e8f0', color: activeTab === 'stats' ? 'white' : '#64748b', boxShadow: activeTab === 'stats' ? '0 4px 6px -1px rgba(30, 58, 138, 0.2)' : 'none', display: 'flex', alignItems: 'center', gap: '8px', transition: 'all 0.2s' }}>
                          <TableIcon size={16} /> Thống kê điểm
                      </button>
                      <button onClick={() => setActiveTab('create')} style={{ padding: '10px 20px', borderRadius: '99px', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 600, background: activeTab === 'create' ? '#1e3a8a' : '#e2e8f0', color: activeTab === 'create' ? 'white' : '#64748b', boxShadow: activeTab === 'create' ? '0 4px 6px -1px rgba(30, 58, 138, 0.2)' : 'none', display: 'flex', alignItems: 'center', gap: '8px', transition: 'all 0.2s' }}>
                          <BrainCircuit size={16} /> Phân tích AI
                      </button>
                  </div>
                  {activeTab === 'stats' && (
                      <div style={{ display: 'flex', gap: '10px' }}>
                          {stats && ( <button onClick={() => exportToExcel('stats-table', fileName)} style={{ padding: '10px 20px', borderRadius: '99px', background: '#2563eb', border: 'none', cursor: 'pointer', fontSize: '14px', fontWeight: 600, color: 'white', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 2px 4px rgba(37, 99, 235, 0.2)' }}> <FileDown size={16} /> Tải file Excel </button> )}
                          <button onClick={() => document.getElementById('re-upload')?.click()} style={{ padding: '10px 20px', borderRadius: '99px', background: 'white', border: '1px solid #cbd5e1', cursor: 'pointer', fontSize: '14px', fontWeight: 600, color: '#475569', display: 'flex', alignItems: 'center', gap: '8px', boxShadow: '0 1px 2px rgba(0,0,0,0.05)' }}> <RefreshCw size={16} /> Tải file khác </button>
                          <input type="file" accept=".xlsx,.csv" onChange={handleDataUpload} style={{ display: 'none' }} id="re-upload" />
                      </div>
                  )}
              </div>

              <div style={{ flex: 1, overflow: 'auto', padding: '0 24px 24px 24px' }}>
                  {activeTab === 'stats' && (
                      <div style={{ animation: 'fadeIn 0.3s ease-out', height: '100%', display: 'flex', flexDirection: 'column' }}>
                          {!stats && (
                              <div style={{ padding: '60px', background: 'white', borderRadius: '16px', border: '2px dashed #cbd5e1', textAlign: 'center', transition: 'border-color 0.2s', maxWidth: '600px', margin: '40px auto' }} onDragOver={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#3b82f6'; }} onDragLeave={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#cbd5e1'; }} onDrop={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#cbd5e1'; }}>
                                  <input type="file" accept=".xlsx,.csv" onChange={handleDataUpload} style={{ display: 'none' }} id="data-upload" />
                                  <label htmlFor="data-upload" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '15px' }}>
                                      <div style={{ width: '80px', height: '80px', background: '#eff6ff', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center' }}> <Upload size={36} color="#3b82f6" /> </div>
                                      <div> <div style={{ fontSize: '18px', fontWeight: 600, color: '#1e293b' }}> Tải lên file ZipGrade </div> <div style={{ fontSize: '14px', color: '#64748b', marginTop: '6px' }}>Hỗ trợ định dạng Excel (.xlsx) hoặc CSV</div> </div>
                                  </label>
                              </div>
                          )}
                          {processedResults && stats && (
                              <div style={{ background: 'white', borderRadius: '12px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', display: 'flex', flexDirection: 'column', height: '100%', border: '1px solid #e2e8f0' }}>
                                  <div style={{ padding: '12px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTopLeftRadius: '12px', borderTopRightRadius: '12px' }}>
                                      <div style={{ display: 'flex', gap: '24px', fontSize: '13px', fontWeight: 500, alignItems: 'center' }}>
                                          <div style={{ display: 'flex', gap: '15px', borderRight: '1px solid #cbd5e1', paddingRight: '15px' }}>
                                              <div style={{ display: 'flex', gap: '6px', alignItems: 'center', padding: '4px 8px', borderRadius: '6px', background: 'white', border: '1px solid #e2e8f0', color: '#1e3a8a', fontWeight: 600 }}> <span style={{width:'12px', height:'12px', background: DEFAULT_COLORS.blue, borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> 0 Sai </div>
                                              <div style={{ display: 'flex', gap: '6px', alignItems: 'center', padding: '4px 8px', borderRadius: '6px', background: 'white', border: '1px solid #e2e8f0', color: '#1e3a8a', fontWeight: 600 }}> <span style={{width:'12px', height:'12px', background: customColors.lowError, borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> &lt;{thresholds.lowCount} Sai </div>
                                              <div style={{ display: 'flex', gap: '6px', alignItems: 'center', padding: '4px 8px', borderRadius: '6px', background: 'white', border: '1px solid #e2e8f0', color: '#1e3a8a', fontWeight: 600 }}> <span style={{width:'12px', height:'12px', background: customColors.highError, borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> &gt;{thresholds.highPercent}% Sai </div>
                                          </div>
                                          <div style={{ display: 'flex', gap: '15px', color: '#334155' }}> <div>TB: <strong>{summaryStats.avg}</strong></div> <div>Max: <strong style={{color:'#16a34a'}}>{summaryStats.max}</strong></div> <div>Min: <strong style={{color:'#dc2626'}}>{summaryStats.min}</strong></div> </div>
                                      </div>
                                      <div style={{ fontSize: '13px', color: '#64748b' }}> File: <strong>{fileName}</strong> </div>
                                  </div>
                                  <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
                                      <table id="stats-table" style={{ width: '100%', fontSize: '12px', borderCollapse: 'separate', borderSpacing: 0, minWidth: '1500px' }}>
                                          <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                                              <tr style={{ background: '#f1f5f9' }}>
                                                  <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', width: '40px' }}>STT</th>
                                                  <th onClick={() => handleSort('sbd')} style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', cursor: 'pointer', userSelect: 'none' }}> <div style={{display:'flex', alignItems:'center', justifyContent:'center', gap:'4px'}}>SBD {renderSortIcon('sbd')}</div> </th>
                                                  <th onClick={() => handleSort('name')} style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', textAlign: 'left', paddingLeft: '10px', cursor: 'pointer', userSelect: 'none', minWidth: '220px' }}> <div style={{display:'flex', alignItems:'center', gap:'4px'}}>Họ và Tên {renderSortIcon('name')}</div> </th>
                                                  <th style={{ borderBottom: '1px solid #cbd5e1' }}>Mã</th>
                                                  <th onClick={() => handleSort('total')} style={{ borderBottom: '1px solid #cbd5e1', cursor: 'pointer', userSelect: 'none' }}> <div style={{display:'flex', alignItems:'center', justifyContent:'center', gap:'4px'}}>Điểm {renderSortIcon('total')}</div> </th>
                                                  <th onClick={() => handleSort('p1')} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '10px', cursor: 'pointer', userSelect: 'none' }}>P1 {renderSortIcon('p1')}</th>
                                                  <th onClick={() => handleSort('p2')} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '10px', cursor: 'pointer', userSelect: 'none' }}>P2 {renderSortIcon('p2')}</th>
                                                  <th onClick={() => handleSort('p3')} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '10px', borderRight: '2px solid #94a3b8', cursor: 'pointer', userSelect: 'none' }}>P3 {renderSortIcon('p3')}</th>
                                                  {stats.map(s => { const isPart2 = s.index >= p2Range.start && s.index <= p2Range.end; return ( <th key={s.index} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '11px', minWidth: '24px', background: isPart2 ? '#fefce8' : '#f1f5f9' }}> {getPart2Label(s.index, activeSubject as any)} </th> ); })}
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
                                                  {stats.map(s => { const isPart2 = s.index >= p2Range.start && s.index <= p2Range.end; return ( <th key={s.index} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '11px', color: '#16a34a', background: isPart2 ? '#e2e8f0' : '#e2e8f0' }}> {s.correctKey} </th> ); })}
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
                                                  {stats.map(s => ( <th key={s.index} style={{ padding: '4px', background: getCellColor(s.wrongCount, s.wrongPercent), borderBottom: '2px solid #94a3b8', minWidth: '24px', fontSize: '10px', color: '#475569' }}> <div>{s.wrongCount}</div> <div style={{fontSize: '9px', opacity: 0.8}}>{s.wrongPercent}%</div> </th> ))}
                                              </tr>
                                          </thead>
                                          <tbody>
                                              {sortedResults.map((st, idx) => (
                                                  <tr key={idx} style={{ background: idx % 2 === 0 ? 'white' : '#fcfcfc' }}>
                                                      <td style={{ position: 'sticky', left: 0, background: idx % 2 === 0 ? 'white' : '#fcfcfc', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9' }}>{idx + 1}</td>
                                                      <td style={{ position: 'sticky', left: '40px', background: idx % 2 === 0 ? 'white' : '#fcfcfc', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9', fontFamily: 'monospace' }}>{st.sbd}</td>
                                                      <td style={{ position: 'sticky', left: '100px', background: idx % 2 === 0 ? 'white' : '#fcfcfc', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9', textAlign: 'left', fontWeight: 600, paddingLeft: '10px', verticalAlign: 'middle', minWidth: '220px' }}> <div style={{ display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden', textOverflow: 'ellipsis', lineHeight: '1.4', maxHeight: '2.8em' }}> {st.name} </div> </td>
                                                      <td style={{ borderBottom: '1px solid #f1f5f9' }}>{st.code}</td>
                                                      <td style={{ borderBottom: '1px solid #f1f5f9', fontWeight: 'bold', color: getScoreColor(st.scores.total) }}>{st.scores.total}</td>
                                                      <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px' }}>{st.scores.p1}</td>
                                                      <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px' }}>{st.scores.p2}</td>
                                                      <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px', borderRight: '2px solid #e2e8f0' }}>{st.scores.p3}</td>
                                                      {stats.map(s => { const isCorrect = st.details[s.index] === 'T'; const isPart2 = s.index >= p2Range.start && s.index <= p2Range.end; let bgColor = 'transparent'; if (isCorrect) { bgColor = isPart2 ? DEFAULT_COLORS.yellow : 'transparent'; } else { bgColor = '#fecaca'; } return ( <td key={s.index} style={{ borderBottom: '1px solid #f1f5f9', background: bgColor, color: isCorrect ? '#cbd5e1' : '#b91c1c', fontSize: '11px', fontWeight: isCorrect ? 400 : 700, padding: '2px', borderLeft: '1px solid #f1f5f9' }}> {isCorrect ? '•' : st.rawAnswers[s.index]} </td> ); })}
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
                      <div style={{ padding: '24px', maxWidth: '800px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                          <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
                              <h3 style={{ marginTop: 0, color: '#1e3a8a', display: 'flex', alignItems: 'center', gap: '10px' }}>
                                  <BrainCircuit size={24} /> Tạo đề ôn tập AI
                              </h3>
                              <p style={{ color: '#64748b', fontSize: '14px', lineHeight: '1.6' }}>
                                  Hệ thống sẽ phân tích các câu sai nhiều (trên {thresholds.highPercent}%) và tạo ra đề ôn tập tương tự để học sinh luyện tập.
                              </p>
                              
                              <div style={{ marginTop: '20px', padding: '16px', background: '#f8fafc', borderRadius: '8px', border: '1px dashed #cbd5e1' }}>
                                  <label style={{ display: 'block', fontWeight: 600, marginBottom: '8px', color: '#334155' }}>1. Tải lên đề gốc (PDF/Word/Text)</label>
                                  <input type="file" accept=".pdf,.docx,.doc,.txt" onChange={handleExamFileUpload} style={{ width: '100%' }} />
                                  {examFile && <div style={{ marginTop: '8px', fontSize: '13px', color: '#16a34a' }}>Đã tải: {examFile.name}</div>}
                              </div>

                              <div style={{ marginTop: '20px' }}>
                                  <label style={{ display: 'block', fontWeight: 600, marginBottom: '8px', color: '#334155' }}>2. Cấu hình & Tạo đề</label>
                                  <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                                      <button 
                                          onClick={generatePracticeExam} 
                                          disabled={!examFile || !stats || isGenerating}
                                          style={{ padding: '10px 20px', background: isGenerating ? '#94a3b8' : '#1e3a8a', color: 'white', border: 'none', borderRadius: '8px', cursor: isGenerating ? 'not-allowed' : 'pointer', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}
                                      >
                                          {isGenerating ? <Loader2 className="spin" size={18} /> : <BrainCircuit size={18} />}
                                          {isGenerating ? 'Đang phân tích & tạo đề...' : 'Tạo đề ôn tập ngay'}
                                      </button>
                                  </div>
                              </div>
                          </div>

                          {generatedExam && (
                              <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
                                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                                      <h3 style={{ margin: 0, color: '#1e3a8a' }}>Kết quả tạo đề</h3>
                                      <button onClick={() => exportExamToWord(generatedExam, `De_On_Tap_${activeSubject}`)} style={{ padding: '8px 16px', background: '#2563eb', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px' }}>
                                          <Download size={16} /> Tải Word
                                      </button>
                                  </div>
                                  <div style={{ padding: '20px', background: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0', whiteSpace: 'pre-wrap', fontFamily: 'serif', fontSize: '14px', lineHeight: '1.6', maxHeight: '500px', overflowY: 'auto' }}>
                                      {generatedExam}
                                  </div>
                              </div>
                          )}
                      </div>
                  )}
              </div>
          </div>
      )}

      {showSettings && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 50 }}>
              <div style={{ background: 'white', padding: '24px', borderRadius: '12px', width: '400px', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.1)' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                      <h3 style={{ margin: 0, color: '#1e3a8a' }}>Cài đặt</h3>
                      <button onClick={() => setShowSettings(false)} style={{ background: 'transparent', border: 'none', cursor: 'pointer', padding: '4px' }}><X size={20} /></button>
                  </div>

                  <div style={{ marginBottom: '20px' }}>
                      <label style={{ display: 'block', fontSize: '13px', fontWeight: 600, marginBottom: '8px', color: '#334155' }}>Ngưỡng cảnh báo sai</label>
                      <div style={{ display: 'flex', gap: '10px' }}>
                          <div style={{ flex: 1 }}>
                              <label style={{ fontSize: '12px', color: '#64748b' }}>Số lượng sai &lt;</label>
                              <input type="number" value={thresholds.lowCount} onChange={(e) => setThresholds({ ...thresholds, lowCount: Number(e.target.value) })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} />
                          </div>
                          <div style={{ flex: 1 }}>
                              <label style={{ fontSize: '12px', color: '#64748b' }}>Phần trăm sai &gt; (%)</label>
                              <input type="number" value={thresholds.highPercent} onChange={(e) => setThresholds({ ...thresholds, highPercent: Number(e.target.value) })} style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #cbd5e1', marginTop: '4px' }} />
                          </div>
                      </div>
                  </div>

                  <button onClick={() => setShowSettings(false)} style={{ width: '100%', padding: '10px', background: '#1e3a8a', color: 'white', borderRadius: '8px', border: 'none', fontWeight: 600, cursor: 'pointer' }}>Đóng</button>
              </div>
          </div>
      )}
      
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