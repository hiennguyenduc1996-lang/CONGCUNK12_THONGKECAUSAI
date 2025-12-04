import React, { useState, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI } from "@google/genai";
import { Upload, FileText, Download, Loader2, Settings, Key, Eye, EyeOff, Calculator, FlaskConical, Languages, BrainCircuit, Table as TableIcon, X, User, School, BookOpen, ChevronRight, LayoutDashboard, FileSpreadsheet, RefreshCw, ArrowUpDown, ArrowUp, ArrowDown, FileDown, Filter } from 'lucide-react';

// Declare libraries
declare const mammoth: any;
declare const XLSX: any;

// --- Types ---

interface StudentResult {
  sbd: string;
  name: string;
  firstName: string; // Added for sorting
  lastName: string;  // Added for sorting
  code: string;
  rawAnswers: Record<string, string>; // Key: Question Index (1, 2, ...), Value: Answer (A, B, C, D)
  scores: {
    total: number;
    p1: number;
    p2: number;
    p3: number;
  };
  details: Record<string, 'T' | 'F'>; // Key: Question Index, Value: T (Correct) or F (Wrong)
}

interface QuestionStat {
  index: number;
  wrongCount: number;
  wrongPercent: number;
  correctKey: string; // Store the correct answer key
}

interface DocFile {
  id: string;
  name: string;
  content: string; // Base64 or Text
  type: 'pdf' | 'text';
}

interface SubjectConfig {
  id: string;
  name: string;
  type: 'math' | 'science' | 'english';
  totalQuestions: number;
  parts: {
    p1: { start: number; end: number; scorePerQ: number };
    p2: { start: number; end: number; scorePerGroup: number }; // Special logic
    p3: { start: number; end: number; scorePerQ: number };
  };
}

interface ThresholdConfig {
  lowCount: number; // e.g., < 5 students wrong
  highPercent: number; // e.g., > 40% students wrong
}

// --- Constants ---

const SUBJECTS_CONFIG: Record<string, SubjectConfig> = {
  math: {
    id: 'math',
    name: 'Toán Học',
    type: 'math',
    totalQuestions: 34,
    parts: {
      p1: { start: 1, end: 12, scorePerQ: 0.25 },
      p2: { start: 13, end: 28, scorePerGroup: 1.0 }, // 4 questions per group
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
      p2: { start: 19, end: 34, scorePerGroup: 1.0 }, // 4 questions per group
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
  }
};

const TABLE_COLORS = ['#dbeafe', '#fef9c3', '#fee2e2']; // Blue, Yellow, Red

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

  // Calculate group number and character
  // Math: 13-16 -> 1a-1d
  const relativeIndex = index - p2.start;
  const groupNum = Math.floor(relativeIndex / 4) + 1;
  const charCode = 97 + (relativeIndex % 4); // 97 is 'a'
  
  return `${groupNum}${String.fromCharCode(charCode)}`;
};

const exportToWord = (elementId: string, fileName: string) => {
    const element = document.getElementById(elementId);
    if (!element) return;

    const htmlContent = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
            <meta charset="utf-8">
            <title>Export HTML to Word Document with Landscape Orientation</title>
            <style>
                @page {
                    size: 29.7cm 21cm;
                    margin: 1cm 1cm 1cm 1cm;
                    mso-page-orientation: landscape;
                }
                body {
                    font-family: 'Times New Roman', serif;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                td, th {
                    border: 1px solid black;
                    padding: 5px;
                    text-align: center;
                    font-size: 10pt;
                }
                .bg-yellow { background-color: #fefce8; }
                .bg-red { background-color: #fecaca; }
                .text-red { color: red; font-weight: bold; }
            </style>
        </head>
        <body>
            ${element.outerHTML}
        </body>
        </html>
    `;

    const blob = new Blob(['\ufeff', htmlContent], {
        type: 'application/msword'
    });
    
    // Create download link
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${fileName || 'Thong_ke'}.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};


// --- Scoring Engine ---

const processData = (data: any[], subjectType: 'math' | 'science' | 'english') => {
  const config = SUBJECTS_CONFIG[subjectType];
  
  const results: StudentResult[] = [];
  const questionStats: Record<number, number> = {}; // Index -> Wrong Count
  const correctKeys: Record<number, string> = {}; // Index -> Correct Key

  // Initialize stats
  for (let i = 1; i <= config.totalQuestions; i++) {
    questionStats[i] = 0;
    correctKeys[i] = ''; 
  }

  // First pass to find correct keys (assuming keys are consistent in ZipGrade CSV or at least present in first row)
  if (data.length > 0) {
    const firstRow = data[0];
    for (let i = 1; i <= config.totalQuestions; i++) {
       const keyCol = `PriKey${i}`;
       if (firstRow[keyCol]) {
         correctKeys[i] = String(firstRow[keyCol]).trim().toUpperCase();
       }
    }
  }

  data.forEach(row => {
    // Basic validation to ensure it's a student row (has StudentID or Name)
    if (!row['StudentID'] && !row['LastName'] && !row['FirstName']) return;

    let p1Score = 0;
    let p2Score = 0;
    let p3Score = 0;
    const details: Record<string, 'T' | 'F'> = {};
    const rawAnswers: Record<string, string> = {};

    // Helper to get answer and key safely from ZipGrade format
    const checkQuestion = (idx: number) => {
      // ZipGrade format: Stu1, Stu2... and PriKey1, PriKey2...
      const stCol = `Stu${idx}`;
      const keyCol = `PriKey${idx}`;
      
      const stAns = String(row[stCol] || '').trim().toUpperCase();
      // Use pre-fetched key or fallback to row key
      const keyAns = correctKeys[idx] || String(row[keyCol] || '').trim().toUpperCase();
      
      // Update global key map if missing
      if (!correctKeys[idx] && keyAns) correctKeys[idx] = keyAns;

      rawAnswers[idx] = stAns;
      
      if (!keyAns) return false;

      const isCorrect = stAns === keyAns;
      return isCorrect;
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

    // --- Part 2 (Group Logic) ---
    if (config.parts.p2.end > 0) {
      // Iterate by groups of 4
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

    // Handle rounding errors
    p1Score = Math.round(p1Score * 100) / 100;
    p2Score = Math.round(p2Score * 100) / 100;
    p3Score = Math.round(p3Score * 100) / 100;
    const totalScore = Math.round((p1Score + p2Score + p3Score) * 100) / 100;

    results.push({
      sbd: String(row['StudentID'] || ''),
      firstName: String(row['FirstName'] || ''),
      lastName: String(row['LastName'] || ''),
      name: `${row['LastName'] || ''} ${row['FirstName'] || ''}`.trim(),
      code: String(row['Key Version'] || row['Exam Code'] || '---'),
      rawAnswers,
      scores: {
        total: totalScore,
        p1: p1Score,
        p2: p2Score,
        p3: p3Score,
      },
      details
    });
  });

  // Calculate percentages
  const stats: QuestionStat[] = [];
  const totalStudents = results.length;
  for (let i = 1; i <= config.totalQuestions; i++) {
    stats.push({
      index: i,
      wrongCount: questionStats[i],
      wrongPercent: totalStudents > 0 ? parseFloat(((questionStats[i] / totalStudents) * 100).toFixed(1)) : 0,
      correctKey: correctKeys[i] || '-'
    });
  }

  return { results, stats };
};


// --- Main App Component ---

const App = () => {
  const [activeSubject, setActiveSubject] = useState<'math' | 'science' | 'english'>('math');
  const [activeTab, setActiveTab] = useState<'stats' | 'create'>('stats');
  
  // Data State
  const [data, setData] = useState<any[] | null>(null);
  const [processedResults, setProcessedResults] = useState<StudentResult[] | null>(null);
  const [stats, setStats] = useState<QuestionStat[] | null>(null);
  const [fileName, setFileName] = useState<string>("");

  // Stats Filter
  const [statsPartFilter, setStatsPartFilter] = useState<'all' | 'p1' | 'p2' | 'p3'>('all');

  // Sorting State
  const [sortConfig, setSortConfig] = useState<{ key: string; direction: 'asc' | 'desc' } | null>(null);

  // Settings & Thresholds
  const [thresholds, setThresholds] = useState<ThresholdConfig>(() => {
    const saved = localStorage.getItem('thresholds');
    return saved ? JSON.parse(saved) : { lowCount: 5, highPercent: 40 };
  });

  // Exam Creation State
  const [examFile, setExamFile] = useState<DocFile | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedExam, setGeneratedExam] = useState<string>("");

  // Settings UI
  const [showSettings, setShowSettings] = useState(false);
  const [userApiKey, setUserApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [showKey, setShowKey] = useState(false);

  useEffect(() => {
    localStorage.setItem('gemini_api_key', userApiKey);
  }, [userApiKey]);

  useEffect(() => {
    localStorage.setItem('thresholds', JSON.stringify(thresholds));
  }, [thresholds]);

  // Handle File Upload (Excel/CSV)
  const handleDataUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const jsonData = XLSX.utils.sheet_to_json(ws);
      setData(jsonData);
      
      // Auto process
      const { results, stats } = processData(jsonData, activeSubject);
      setProcessedResults(results);
      setStats(stats);
      setSortConfig(null); // Reset sort on new file
    };
    reader.readAsBinaryString(file);
    e.target.value = ''; // Reset
  };

  // Re-process when subject changes
  useEffect(() => {
    if (data) {
      const { results, stats: newStats } = processData(data, activeSubject);
      setProcessedResults(results);
      setStats(newStats);
    }
  }, [activeSubject]);

  // Handle Sort
  const handleSort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  // Get Sorted Data
  const sortedResults = useMemo(() => {
    if (!processedResults) return [];
    if (!sortConfig) return processedResults;

    const sorted = [...processedResults];
    sorted.sort((a, b) => {
      let aVal: any = '';
      let bVal: any = '';

      // Determine values based on key
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

  // Calculate General Stats
  const summaryStats = useMemo(() => {
    if (!processedResults || processedResults.length === 0) return { min: 0, max: 0, avg: 0 };
    const scores = processedResults.map(r => r.scores.total);
    const min = Math.min(...scores);
    const max = Math.max(...scores);
    const sum = scores.reduce((a, b) => a + b, 0);
    const avg = parseFloat((sum / scores.length).toFixed(2));
    return { min, max, avg };
  }, [processedResults]);

  // Filter Wrong Stats based on user selection
  const filteredWrongStats = useMemo(() => {
      if (!stats) return [];
      const config = SUBJECTS_CONFIG[activeSubject];
      
      return stats.filter(s => {
          // Check percentage threshold first
          if (s.wrongPercent < thresholds.highPercent) return false;

          // Check Part Filter
          if (statsPartFilter === 'all') return true;
          
          const idx = s.index;
          if (statsPartFilter === 'p1') {
              return idx >= config.parts.p1.start && idx <= config.parts.p1.end;
          }
          if (statsPartFilter === 'p2') {
              return idx >= config.parts.p2.start && idx <= config.parts.p2.end;
          }
          if (statsPartFilter === 'p3') {
              return idx >= config.parts.p3.start && idx <= config.parts.p3.end;
          }
          return false;
      }).sort((a,b) => b.wrongCount - a.wrongCount);
  }, [stats, statsPartFilter, thresholds.highPercent, activeSubject]);

  // Handle Exam File Upload
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
      const ai = new GoogleGenAI({ apiKey: userApiKey || process.env.API_KEY || '' });
      
      // Use filtered stats
      const topWrong = filteredWrongStats.slice(0, 5); // Take top 5 of filtered
      const wrongIndices = topWrong.map(s => getPart2Label(s.index, activeSubject)).join(", ");
      
      const prompt = `
        Bạn là một giáo viên chuyên nghiệp. Dưới đây là nội dung của một đề thi gốc và thống kê các câu hỏi mà học sinh làm sai nhiều nhất (Lọc theo tiêu chí: ${statsPartFilter === 'all' ? 'Tất cả' : statsPartFilter.toUpperCase()}, Tỷ lệ sai > ${thresholds.highPercent}%).
        
        Nhiệm vụ:
        1. Phân tích nội dung kiến thức của các câu hỏi bị sai nhiều (Câu số: ${wrongIndices}).
        2. Tạo ra một đề ôn tập ngắn (khoảng 5-10 câu) tập trung vào các dạng bài/kiến thức đó để giúp học sinh khắc phục lỗi sai.
        3. Đề ôn tập cần có đáp án và lời giải chi tiết ở cuối.
        
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

  // Render Stats Cell Color
  const getCellColor = (wrongCount: number, wrongPercent: number) => {
    if (wrongCount === 0) return TABLE_COLORS[0]; // Blue
    if (wrongPercent > thresholds.highPercent) return TABLE_COLORS[2]; // Red
    if (wrongCount < thresholds.lowCount) return TABLE_COLORS[1]; // Yellow
    return 'white';
  };

  const getScoreColor = (score: number) => {
    if (score >= 8) return '#16a34a'; // Green
    if (score >= 5) return '#ca8a04'; // Yellow
    return '#dc2626'; // Red
  };

  const renderSortIcon = (key: string) => {
    if (!sortConfig || sortConfig.key !== key) return <ArrowUpDown size={12} style={{ opacity: 0.3 }} />;
    return sortConfig.direction === 'asc' ? <ArrowUp size={12} /> : <ArrowDown size={12} />;
  };

  // --- UI RENDER ---

  // Settings Overlay
  if (showSettings) {
    return (
      <div style={{ position: 'fixed', inset: 0, background: 'rgba(15, 23, 42, 0.6)', backdropFilter: 'blur(4px)', zIndex: 2000, display: 'flex', justifyContent: 'center', alignItems: 'center' }}>
         <div style={{ background: 'white', borderRadius: '24px', boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)', width: '600px', overflow: 'hidden' }}>
             
             {/* Header */}
             <div style={{ background: '#1e3a8a', padding: '20px 30px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                 <h2 style={{ margin: 0, color: 'white', fontSize: '18px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <Settings size={20} /> Cài đặt hệ thống
                 </h2>
                 <button onClick={() => setShowSettings(false)} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: 'white', borderRadius: '50%', padding: '6px', cursor: 'pointer', display: 'flex' }}>
                    <X size={18} />
                 </button>
             </div>

             <div style={{ padding: '30px', maxHeight: '70vh', overflowY: 'auto' }}>
                 
                 {/* Author Info */}
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

                 {/* Threshold Configuration */}
                 <div style={{ marginBottom: '30px' }}>
                    <h4 style={{ margin: '0 0 15px 0', color: '#334155' }}>Cấu hình hiển thị thống kê</h4>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                       <div>
                          <label style={{ display: 'block', marginBottom: '8px', fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                             Số lượng sai ít (Tô vàng)
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
                             Tỷ lệ sai nhiều (Tô đỏ)
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
                 </div>

                 {/* API Key */}
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

             {/* Footer */}
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

  const p2Range = SUBJECTS_CONFIG[activeSubject].parts.p2;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', overflow: 'hidden', background: '#f8fafc' }}>
      
      {/* --- HEADER --- */}
      <header style={{ height: '64px', background: '#1e3a8a', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '0 24px', color: 'white', flexShrink: 0 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', fontWeight: 700, fontSize: '18px' }}>
              <div style={{ width: '36px', height: '36px', background: 'white', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#1e3a8a' }}>
                  <FileSpreadsheet size={20} />
              </div>
              <div>
                  <div style={{ lineHeight: '1.2' }}>Converter</div>
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

      {/* --- SUBJECT NAVIGATION BAR --- */}
      <div style={{ background: 'white', borderBottom: '1px solid #e2e8f0', padding: '12px 24px', display: 'flex', gap: '10px', overflowX: 'auto' }}>
          {Object.values(SUBJECTS_CONFIG).map(subj => {
              const isActive = activeSubject === subj.id;
              return (
                <button
                  key={subj.id}
                  onClick={() => setActiveSubject(subj.type)}
                  style={{
                      display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 20px', 
                      borderRadius: '99px', // Pill shape
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
                  <span>{subj.name}</span>
                </button>
              );
          })}
      </div>

      {/* --- MAIN CONTENT --- */}
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
          
          {/* Sub-Header / Toolbar */}
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
                           onClick={() => exportToWord('stats-table', fileName)}
                           style={{
                              padding: '8px 16px', background: '#2563eb', border: 'none', borderRadius: '8px',
                              cursor: 'pointer', fontSize: '13px', fontWeight: 600, color: 'white',
                              display: 'flex', alignItems: 'center', gap: '8px',
                              boxShadow: '0 2px 4px rgba(37, 99, 235, 0.2)'
                           }}
                        >
                           <FileDown size={14} /> Tải file Word
                        </button>
                     )}
                     <button
                        onClick={() => document.getElementById('re-upload')?.click()}
                        style={{
                           padding: '8px 16px', background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px',
                           cursor: 'pointer', fontSize: '13px', fontWeight: 600, color: '#475569',
                           display: 'flex', alignItems: 'center', gap: '8px'
                        }}
                     >
                        <RefreshCw size={14} /> Tải file khác
                     </button>
                     <input type="file" accept=".xlsx,.csv" onChange={handleDataUpload} style={{ display: 'none' }} id="re-upload" />
                  </div>
              )}
          </div>

          {/* Workspace */}
          <div style={{ flex: 1, overflow: 'auto', padding: '0 24px 24px 24px' }}>
             
                {activeTab === 'stats' && (
                    <div style={{ animation: 'fadeIn 0.3s ease-out', height: '100%', display: 'flex', flexDirection: 'column' }}>
                        
                        {/* Empty State / Upload */}
                        {!stats && (
                           <div style={{ 
                                padding: '60px', background: 'white', borderRadius: '16px', border: '2px dashed #cbd5e1', 
                                textAlign: 'center', transition: 'border-color 0.2s', maxWidth: '600px', margin: '40px auto'
                            }}
                            onDragOver={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#3b82f6'; }}
                            onDragLeave={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#cbd5e1'; }}
                            onDrop={(e) => { e.preventDefault(); e.currentTarget.style.borderColor = '#cbd5e1'; /* Handle drop */ }}
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
                                
                                {/* Top Summary Bar */}
                                <div style={{ padding: '12px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#f8fafc', borderTopLeftRadius: '12px', borderTopRightRadius: '12px' }}>
                                    <div style={{ display: 'flex', gap: '24px', fontSize: '13px', fontWeight: 500, alignItems: 'center' }}>
                                        {/* Legend */}
                                        <div style={{ display: 'flex', gap: '15px', borderRight: '1px solid #cbd5e1', paddingRight: '15px' }}>
                                            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}><span style={{width:'10px', height:'10px', background: TABLE_COLORS[0], borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> 0 Sai</div>
                                            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}><span style={{width:'10px', height:'10px', background: TABLE_COLORS[1], borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> &lt;{thresholds.lowCount} Sai</div>
                                            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}><span style={{width:'10px', height:'10px', background: TABLE_COLORS[2], borderRadius:'2px', border:'1px solid #cbd5e1'}}></span> &gt;{thresholds.highPercent}% Sai</div>
                                        </div>
                                        {/* Stats */}
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

                                {/* Scrollable Table Area */}
                                <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
                                    <table id="stats-table" style={{ width: '100%', fontSize: '12px', borderCollapse: 'separate', borderSpacing: 0, minWidth: '1500px' }}>
                                        <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                                            {/* Row 1: Headers & Sorting */}
                                            <tr style={{ background: '#f1f5f9' }}>
                                                <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', width: '40px' }}>STT</th>
                                                <th onClick={() => handleSort('sbd')} style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', cursor: 'pointer', userSelect: 'none' }}>
                                                   <div style={{display:'flex', alignItems:'center', justifyContent:'center', gap:'4px'}}>SBD {renderSortIcon('sbd')}</div>
                                                </th>
                                                <th onClick={() => handleSort('name')} style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#f1f5f9', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', textAlign: 'left', paddingLeft: '10px', cursor: 'pointer', userSelect: 'none' }}>
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
                                                   const isPart2 = s.index >= p2Range.start && s.index <= p2Range.end;
                                                   return (
                                                    <th key={s.index} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '11px', minWidth: '24px', background: isPart2 ? '#fefce8' : '#f1f5f9' }}>
                                                        {getPart2Label(s.index, activeSubject)}
                                                    </th>
                                                   );
                                                })}
                                            </tr>
                                            
                                            {/* Row 2: Correct Keys */}
                                            <tr style={{ background: '#e2e8f0' }}>
                                                <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#e2e8f0', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1' }}></th>
                                                <th style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#e2e8f0', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1' }}></th>
                                                <th style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#e2e8f0', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', textAlign: 'left', paddingLeft: '10px', color: '#475569', fontSize: '11px' }}>Đáp án đúng</th>
                                                <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                <th style={{ borderBottom: '1px solid #cbd5e1' }}></th>
                                                <th style={{ borderBottom: '1px solid #cbd5e1', borderRight: '2px solid #94a3b8' }}></th>
                                                {stats.map(s => {
                                                   const isPart2 = s.index >= p2Range.start && s.index <= p2Range.end;
                                                   return (
                                                    <th key={s.index} style={{ borderBottom: '1px solid #cbd5e1', fontSize: '11px', color: '#16a34a', background: isPart2 ? '#fefce8' : '#e2e8f0' }}>
                                                        {s.correctKey}
                                                    </th>
                                                   );
                                                })}
                                            </tr>

                                            {/* Row 3: Stats */}
                                            <tr style={{ background: '#f8fafc' }}>
                                                <th style={{ position: 'sticky', left: 0, zIndex: 11, background: '#f8fafc', borderRight: '1px solid #cbd5e1', borderBottom: '2px solid #94a3b8' }}></th>
                                                <th style={{ position: 'sticky', left: '40px', zIndex: 11, background: '#f8fafc', borderRight: '1px solid #cbd5e1', borderBottom: '2px solid #94a3b8' }}></th>
                                                <th style={{ position: 'sticky', left: '100px', zIndex: 11, background: '#f8fafc', borderRight: '1px solid #cbd5e1', borderBottom: '2px solid #94a3b8', textAlign: 'left', paddingLeft: '10px', color: '#64748b', fontSize: '11px' }}>Thống kê (Số lượng/%)</th>
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
                                                    <td style={{ position: 'sticky', left: '100px', background: idx % 2 === 0 ? 'white' : '#fcfcfc', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #f1f5f9', textAlign: 'left', fontWeight: 600, paddingLeft: '10px' }}>{st.name}</td>
                                                    <td style={{ borderBottom: '1px solid #f1f5f9' }}>{st.code}</td>
                                                    <td style={{ borderBottom: '1px solid #f1f5f9', fontWeight: 'bold', color: getScoreColor(st.scores.total) }}>{st.scores.total}</td>
                                                    <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px' }}>{st.scores.p1}</td>
                                                    <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px' }}>{st.scores.p2}</td>
                                                    <td style={{ borderBottom: '1px solid #f1f5f9', color: '#64748b', fontSize: '11px', borderRight: '2px solid #e2e8f0' }}>{st.scores.p3}</td>
                                                    {stats.map(s => {
                                                        const isCorrect = st.details[s.index] === 'T';
                                                        const isPart2 = s.index >= p2Range.start && s.index <= p2Range.end;
                                                        const wrongBg = '#fecaca'; // Red 200 - Darker Red for cells
                                                        const part2CorrectBg = '#fefce8'; // Light Yellow

                                                        // Add explicit class for Word Export to recognize styles
                                                        const cellClass = isCorrect ? (isPart2 ? 'bg-yellow' : '') : 'bg-red text-red';

                                                        return (
                                                            <td key={s.index} className={cellClass} style={{ 
                                                                borderBottom: '1px solid #f1f5f9', 
                                                                background: isCorrect ? (isPart2 ? part2CorrectBg : 'transparent') : wrongBg,
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
                    <div style={{ display: 'grid', gridTemplateColumns: '350px 1fr', gap: '30px', animation: 'fadeIn 0.3s ease-out', maxWidth: '1400px', margin: '0 auto' }}>
                        {/* Control Panel */}
                        <div>
                            <div style={{ background: 'white', padding: '25px', borderRadius: '16px', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)', border: '1px solid #e2e8f0' }}>
                                <h3 style={{ margin: '0 0 20px 0', fontSize: '16px', color: '#1e3a8a', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <BookOpen size={18} /> Nguồn đề thi
                                </h3>
                                
                                <div style={{ 
                                    padding: '30px 20px', background: '#f8fafc', borderRadius: '12px', border: '2px dashed #e2e8f0', 
                                    textAlign: 'center', marginBottom: '20px', cursor: 'pointer', transition: 'all 0.2s'
                                }}
                                onClick={() => document.getElementById('exam-upload')?.click()}
                                onMouseOver={(e) => e.currentTarget.style.borderColor = '#93c5fd'}
                                onMouseOut={(e) => e.currentTarget.style.borderColor = '#e2e8f0'}
                                >
                                    <input type="file" accept=".pdf,.docx,.doc" onChange={handleExamFileUpload} style={{ display: 'none' }} id="exam-upload" />
                                    <div style={{ width: '40px', height: '40px', background: 'white', borderRadius: '50%', margin: '0 auto 10px auto', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                                        <Upload size={20} color="#64748b" />
                                    </div>
                                    <div style={{ fontSize: '14px', fontWeight: 600, color: '#334155' }}>
                                        {examFile ? examFile.name : "Chọn file đề gốc"}
                                    </div>
                                    <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '4px' }}>Hỗ trợ PDF, DOCX</div>
                                </div>

                                {stats && (
                                    <div style={{ marginBottom: '20px' }}>
                                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                                            <div style={{ fontSize: '13px', fontWeight: 600, color: '#475569' }}>
                                                Thống kê lỗi sai (&gt;{thresholds.highPercent}%):
                                            </div>
                                            {/* Filter Dropdown */}
                                            <div style={{ position: 'relative' }}>
                                                <select 
                                                    value={statsPartFilter} 
                                                    onChange={(e) => setStatsPartFilter(e.target.value as any)}
                                                    style={{ 
                                                        padding: '4px 8px', borderRadius: '6px', border: '1px solid #cbd5e1', 
                                                        fontSize: '11px', color: '#334155', cursor: 'pointer', outline: 'none' 
                                                    }}
                                                >
                                                    <option value="all">Tất cả</option>
                                                    <option value="p1">Phần 1</option>
                                                    <option value="p2">Phần 2</option>
                                                    <option value="p3">Phần 3</option>
                                                </select>
                                            </div>
                                        </div>
                                        
                                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                                            {filteredWrongStats.length > 0 ? (
                                                filteredWrongStats.map(s => (
                                                    <div key={s.index} style={{ fontSize: '12px', padding: '4px 8px', background: '#fee2e2', color: '#b91c1c', borderRadius: '6px', fontWeight: 600 }}>
                                                        Câu {getPart2Label(s.index, activeSubject)} ({s.wrongPercent}%)
                                                    </div>
                                                ))
                                            ) : (
                                                <div style={{ fontSize: '12px', color: '#94a3b8', fontStyle: 'italic' }}>Không có câu nào thỏa mãn điều kiện lọc.</div>
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
                                        transition: 'transform 0.1s'
                                    }}
                                    onMouseDown={(e) => !isGenerating && (e.currentTarget.style.transform = 'scale(0.98)')}
                                    onMouseUp={(e) => !isGenerating && (e.currentTarget.style.transform = 'scale(1)')}
                                >
                                    {isGenerating ? <Loader2 className="spin" size={20} /> : <BrainCircuit size={20} />} 
                                    {isGenerating ? 'Đang phân tích...' : 'Tạo đề ôn tập'}
                                </button>
                            </div>
                        </div>

                        {/* Result View */}
                        <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                            <div style={{ flex: 1, background: 'white', borderRadius: '16px', border: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', overflow: 'hidden', boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)' }}>
                                <div style={{ padding: '15px 20px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <div style={{ fontWeight: 600, color: '#334155', fontSize: '14px' }}>Nội dung đề tạo bởi AI</div>
                                    {generatedExam && (
                                        <button style={{ padding: '6px 12px', background: '#16a34a', color: 'white', border: 'none', borderRadius: '6px', fontSize: '12px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                            <Download size={14} /> Sao chép / Tải về
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

             </div>
      </div>
      
      <style>{`
        .spin { animation: spin 1s linear infinite; }
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        
        /* Custom scrollbar for main area */
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
