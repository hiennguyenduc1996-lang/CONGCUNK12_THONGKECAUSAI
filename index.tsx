import React, { useState, useEffect, useMemo, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI } from "@google/genai";
import { Upload, FileText, Download, Loader2, Settings, Key, Eye, EyeOff, Calculator, FlaskConical, Languages, BrainCircuit, Table as TableIcon, X, User, School, BookOpen, ChevronRight, LayoutDashboard, FileSpreadsheet, RefreshCw, ArrowUpDown, ArrowUp, ArrowDown, FileDown, Filter, Palette, Monitor, Hourglass, TrendingUp, Users, Database, Sigma, Award, Trash2, Atom, Globe, ScrollText, CheckSquare, Square, Cloud, Share2, Copy, ExternalLink, HelpCircle, Save, Link, ArrowRight, Laptop, AlertTriangle } from 'lucide-react';

// Declare libraries
declare const mammoth: any;
declare const XLSX: any;

// --- Types ---

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

const exportToExcel = (elementId: string, fileName: string) => {
    const table = document.getElementById(elementId);
    if (!table || typeof XLSX === 'undefined') return;
    const wb = XLSX.utils.table_to_book(table, { sheet: "ThongKe" });
    XLSX.writeFile(wb, `${fileName || 'Thong_ke'}.xlsx`);
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


const getBlockType = (className: string): 'A' | 'A1' | 'B' | 'Other' => {
    const c = String(className || '').toUpperCase();
    if (c.includes('E')) return 'A1';
    if (c.includes('B')) return 'B';
    if (c.includes('A')) return 'A';
    return 'Other';
};

// --- TOP-LEVEL RESPONSIVE CHART COMPONENT ---
const ResponsiveSVGChart = ({
    students,
    chartClassStudents,
    activeExams,
    rankingMap,
    visibleStudents,
    hoveredStudentId,
    setHoveredStudentId,
    useFullScale,
    comparisonMode,
    selectedChartClasses,
    getStudentColor,
    showAllLabels
}: {
    students: StudentProfile[];
    chartClassStudents: StudentProfile[];
    activeExams: number[];
    rankingMap: Record<number, Record<string, { rank: number; total: number; score: number }>>;
    visibleStudents: Record<string, boolean>;
    hoveredStudentId: string | null;
    setHoveredStudentId: (id: string | null) => void;
    useFullScale: boolean;
    comparisonMode: 'class' | 'block' | 'school';
    selectedChartClasses: string[];
    getStudentColor: (idx: number) => string;
    showAllLabels: boolean;
}) => {
    const containerRef = useRef<HTMLDivElement>(null);
    const [dimensions, setDimensions] = useState({ width: 600, height: 380 });
    const [tooltip, setTooltip] = useState<{ x: number, y: number, studentName: string, exam: number, score: number, rank: number, total: number, color: string } | null>(null);

    useEffect(() => {
        if (!containerRef.current) return;
        const resizeObserver = new ResizeObserver((entries) => {
            if (!entries || entries.length === 0) return;
            const { width, height } = entries[0].contentRect;
            const finalHeight = height > 0 ? height : 380;
            setDimensions((prev) => {
                const newW = Math.max(width, 300);
                const newH = Math.max(finalHeight, 350);
                // Only update dimensions if there is a significant change (> 5px)
                // This breaks the hover layout loop and completely stops the jumping/flickering bug
                if (Math.abs(prev.width - newW) > 5 || Math.abs(prev.height - newH) > 5) {
                    return { width: newW, height: newH };
                }
                return prev;
            });
        });
        resizeObserver.observe(containerRef.current);
        return () => resizeObserver.disconnect();
    }, []);

    const { width, height } = dimensions;
    const paddingLeft = 70;
    const paddingRight = 40;
    const paddingTop = 50;
    const paddingBottom = 45;

    const chartW = width - paddingLeft - paddingRight;
    const chartH = height - paddingTop - paddingBottom;

    const comparisonGroupSize = useMemo(() => {
        if (comparisonMode === 'class') {
            let maxCount = 0;
            selectedChartClasses.forEach(c => {
                const count = students.filter(s => s && s.class === c).length;
                if (count > maxCount) maxCount = count;
            });
            return maxCount || students.filter(s => s && selectedChartClasses.includes(s.class)).length;
        } else if (comparisonMode === 'block') {
            const currentBlocks = Array.from(new Set(selectedChartClasses.map(c => getBlockType(c))));
            return students.filter(s => s && currentBlocks.includes(getBlockType(s.class))).length;
        } else {
            return students.length;
        }
    }, [students, selectedChartClasses, comparisonMode]);

    const maxRankValue = useMemo(() => {
        if (useFullScale) {
            return Math.max(comparisonGroupSize, 2);
        }
        let maxR = 1;
        chartClassStudents.forEach(s => {
            if (!visibleStudents[s.id]) return;
            activeExams.forEach(ex => {
                const r = rankingMap[ex]?.[s.id]?.rank;
                if (r !== undefined && r > maxR) {
                    maxR = r;
                }
            });
        });
        return Math.max(maxR + 1, 10);
    }, [useFullScale, comparisonGroupSize, chartClassStudents, visibleStudents, activeExams, rankingMap]);

    const getX = (index: number) => {
        if (activeExams.length <= 1) return paddingLeft + chartW / 2;
        return paddingLeft + (index / (activeExams.length - 1)) * chartW;
    };

    const getY = (rank: number) => {
        const range = maxRankValue - 1 || 1;
        return paddingTop + ((rank - 1) / range) * chartH;
    };

    const ticks = useMemo(() => {
        const t = [1];
        let step = 5;
        if (maxRankValue > 100) step = 20;
        else if (maxRankValue > 50) step = 10;
        
        let next = step;
        while (next < maxRankValue) {
            t.push(next);
            next += step;
        }
        if (t[t.length - 1] !== maxRankValue && maxRankValue - t[t.length - 1] >= 2) {
            t.push(maxRankValue);
        }
        return t;
    }, [maxRankValue]);

    const numVisible = useMemo(() => {
        return Object.values(visibleStudents).filter(Boolean).length;
    }, [visibleStudents]);

    return (
        <div ref={containerRef} style={{ width: '100%', height: '100%', position: 'relative' }}>
            <svg width={width} height={height} style={{ overflow: 'visible' }}>
                <defs>
                    <marker
                        id="arrow-up"
                        viewBox="0 0 10 10"
                        refX="5"
                        refY="2"
                        markerWidth="6"
                        markerHeight="6"
                        orient="auto"
                    >
                        <path d="M 0 10 L 5 0 L 10 10 z" fill="#1e293b" />
                    </marker>
                    <marker
                        id="arrow-right"
                        viewBox="0 0 10 10"
                        refX="8"
                        refY="5"
                        markerWidth="6"
                        markerHeight="6"
                        orient="auto"
                    >
                        <path d="M 0 0 L 10 5 L 0 10 z" fill="#1e293b" />
                    </marker>
                </defs>

                {/* Grid Lines & Y Axis Labels */}
                {ticks.map((rValue) => {
                    const y = getY(rValue);
                    return (
                        <g key={rValue}>
                            <line 
                                x1={paddingLeft} 
                                y1={y} 
                                x2={paddingLeft + chartW} 
                                y2={y} 
                                stroke="#f1f5f9" 
                                strokeWidth={1.5}
                                strokeDasharray="4 4" 
                            />
                            <text 
                                x={paddingLeft - 12} 
                                y={y + 4} 
                                textAnchor="end" 
                                style={{ fontSize: '12px', fill: '#475569', fontFamily: 'monospace', fontWeight: 600 }}
                            >
                                {rValue}
                            </text>
                        </g>
                    );
                })}

                {/* X Axis Grid Lines & Exam Labels */}
                {activeExams.map((ex, idx) => {
                    const x = getX(idx);
                    return (
                        <g key={idx}>
                            <line 
                                x1={x} 
                                y1={paddingTop} 
                                x2={x} 
                                y2={paddingTop + chartH} 
                                stroke="#f1f5f9" 
                                strokeWidth={1.5}
                                strokeDasharray="4 4" 
                            />
                            <text 
                                x={x} 
                                y={paddingTop + chartH + 22} 
                                textAnchor="middle" 
                                style={{ fontSize: '12px', fill: '#475569', fontWeight: 700 }}
                            >
                                KT {ex}
                            </text>
                        </g>
                    );
                })}

                {/* Draw Solid Axis Lines */}
                <line 
                    x1={paddingLeft} 
                    y1={paddingTop + chartH} 
                    x2={paddingLeft} 
                    y2={paddingTop - 25} 
                    stroke="#1e293b" 
                    strokeWidth={2} 
                    markerEnd="url(#arrow-up)" 
                />
                <line 
                    x1={paddingLeft} 
                    y1={paddingTop + chartH} 
                    x2={paddingLeft + chartW + 30} 
                    y2={paddingTop + chartH} 
                    stroke="#1e293b" 
                    strokeWidth={2} 
                    markerEnd="url(#arrow-right)" 
                />

                {/* Badge at the top of Y Axis */}
                <g transform={`translate(${paddingLeft - 32}, ${paddingTop - 45})`}>
                    <rect width="64" height="22" rx="11" fill="#1e3a8a" />
                    <text x="32" y="15" textAnchor="middle" style={{ fill: '#ffffff', fontSize: '11px', fontWeight: 700 }}>
                        Xếp hạng
                    </text>
                </g>

                {/* Student Lines and Points */}
                {chartClassStudents.map((s, sIdx) => {
                    const isVisible = !!visibleStudents[s.id];
                    if (!isVisible) return null;

                    const color = getStudentColor(sIdx);
                    const isHovered = hoveredStudentId === s.id;
                    const isAnyHovered = hoveredStudentId !== null;

                    const points: { x: number; y: number; ex: number; rank: number; score: number; total: number }[] = [];
                    activeExams.forEach((ex, idx) => {
                        const data = rankingMap[ex]?.[s.id];
                        if (data !== undefined) {
                            points.push({
                                x: getX(idx),
                                y: getY(data.rank),
                                ex,
                                rank: data.rank,
                                score: data.score,
                                total: data.total
                            });
                        }
                    });

                    if (points.length === 0) return null;

                    let pathD = '';
                    points.forEach((pt, idx) => {
                        if (idx === 0) pathD += `M ${pt.x} ${pt.y}`;
                        else pathD += ` L ${pt.x} ${pt.y}`;
                    });

                    const opacity = isHovered ? 1 : (isAnyHovered ? 0.45 : 0.8);
                    const strokeWidth = isHovered ? 4.5 : 2.5;

                    return (
                        <g key={s.id}>
                            {/* The line connecting the dots */}
                            <path 
                                d={pathD} 
                                fill="none" 
                                stroke={color} 
                                strokeWidth={strokeWidth} 
                                strokeOpacity={opacity}
                                strokeLinecap="round"
                                strokeLinejoin="round"
                            />

                            {/* Text labels directly above/below circles */}
                            {points.map((pt, idx) => {
                                const showText = showAllLabels || isHovered || numVisible <= 5;
                                if (!showText) return null;

                                const textY = pt.y - 12 < paddingTop ? pt.y + 16 : pt.y - 10;
                                return (
                                    <text 
                                        key={`lbl-${idx}`}
                                        x={pt.x}
                                        y={textY}
                                        textAnchor="middle"
                                        opacity={opacity}
                                        style={{ 
                                            fontSize: isHovered ? '13px' : '11px', 
                                            fontWeight: 'bold', 
                                            fill: color, 
                                            pointerEvents: 'none'
                                        }}
                                    >
                                        {pt.rank}
                                    </text>
                                );
                            })}

                            {/* Interactive Circles */}
                            {points.map((pt, idx) => (
                                <circle 
                                    key={`c-${idx}`}
                                    cx={pt.x}
                                    cy={pt.y}
                                    r={isHovered ? 6 : 4.5}
                                    fill={color}
                                    stroke="white"
                                    strokeWidth={2}
                                    opacity={opacity}
                                    style={{ cursor: 'pointer', transition: 'all 0.15s ease' }}
                                    onMouseEnter={() => {
                                        setHoveredStudentId(s.id);
                                        setTooltip({
                                            x: pt.x,
                                            y: pt.y,
                                            studentName: s.fullName,
                                            exam: pt.ex,
                                            score: pt.score,
                                            rank: pt.rank,
                                            total: pt.total,
                                            color
                                        });
                                    }}
                                    onMouseLeave={() => {
                                        setHoveredStudentId(null);
                                        setTooltip(null);
                                    }}
                                />
                            ))}
                        </g>
                    );
                })}
            </svg>

            {tooltip && (() => {
                const tooltipWidth = 170;
                const isTooFarRight = tooltip.x + tooltipWidth + 12 > width;
                const tooltipLeft = isTooFarRight ? tooltip.x - tooltipWidth - 12 : tooltip.x + 12;
                const tooltipTop = tooltip.y - 50 < 10 ? tooltip.y + 15 : tooltip.y - 50;
                return (
                    <div 
                        style={{
                            position: 'absolute',
                            left: `${tooltipLeft}px`,
                            top: `${tooltipTop}px`,
                            background: 'rgba(15, 23, 42, 0.95)',
                            color: 'white',
                            padding: '10px 14px',
                            borderRadius: '8px',
                            boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.3)',
                            fontSize: '12px',
                            zIndex: 1000,
                            pointerEvents: 'none',
                            borderLeft: `4px solid ${tooltip.color}`,
                            minWidth: `${tooltipWidth}px`
                        }}
                    >
                        <div style={{ fontWeight: 700, marginBottom: '4px', fontSize: '13px' }}>{tooltip.studentName}</div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', gap: '15px', color: '#cbd5e1', marginBottom: '2px' }}>
                            <span>Lần thi:</span>
                            <span style={{ fontWeight: 600, color: 'white' }}>KT {tooltip.exam}</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', gap: '15px', color: '#cbd5e1', marginBottom: '2px' }}>
                            <span>Điểm số:</span>
                            <span style={{ fontWeight: 600, color: '#f59e0b' }}>{tooltip.score}đ</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', gap: '15px', color: '#cbd5e1' }}>
                            <span>Xếp hạng:</span>
                            <span style={{ fontWeight: 700, color: '#10b981' }}>
                                Hạng {tooltip.rank}/{tooltip.total}
                            </span>
                        </div>
                    </div>
                );
            })()}
        </div>
    );
};

// --- RANKING & SUMMARY COMPONENT ---

const RankingView = () => {
    const [subTab, setSubTab] = useState<'students' | 'scores' | 'summary' | 'sort-summary' | 'cloud' | 'chart'>('students');
    const [students, setStudents] = useState<StudentProfile[]>([]);
    const [examData, setExamData] = useState<ExamDataStore>({});
    const [activeExamTime, setActiveExamTime] = useState<number>(1);
    const [summaryTab, setSummaryTab] = useState<'math'|'phys'|'chem'|'eng'|'bio'|'custom'|'A'|'A1'|'B'|'total'>('math');
    const [sortConfig, setSortConfig] = useState<{ key: string, direction: 'asc'|'desc' } | null>(null);
    
    // Custom combination subjects
    const [customSubjects, setCustomSubjects] = useState({ math: false, phys: true, chem: false, eng: true });

    // Chart tab states
    const [selectedChartClasses, setSelectedChartClasses] = useState<string[]>([]);
    const [isChartClassFilterOpen, setIsChartClassFilterOpen] = useState(false);
    const chartClassFilterRef = useRef<HTMLDivElement>(null);
    const [comparisonMode, setComparisonMode] = useState<'class' | 'block' | 'school'>('class');
    const [rankScoreType, setRankScoreType] = useState<'math' | 'phys' | 'chem' | 'eng' | 'bio' | 'A' | 'A1' | 'B' | 'total'>('total');
    const [visibleStudents, setVisibleStudents] = useState<Record<string, boolean>>({});
    const [hoveredStudentId, setHoveredStudentId] = useState<string | null>(null);
    const [chartSearch, setChartSearch] = useState('');
    const [useFullScale, setUseFullScale] = useState<boolean>(true);
    const [showAllLabels, setShowAllLabels] = useState<boolean>(true);
    const [chartHeight, setChartHeight] = useState<number>(380);
    
    // Cloud State
    const DEFAULT_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzFwQd_LBj4GrNG5P-BlmHKBO4XLrhx6aN0dutYYbTSK2GJRL04J1OSLFc09Gcc3Qlt/exec";
    const [scriptUrl, setScriptUrl] = useState(() => localStorage.getItem('gap_script_url') || DEFAULT_SCRIPT_URL);
    const [syncStatus, setSyncStatus] = useState<'idle' | 'loading' | 'success' | 'error'>('idle');
    const [syncMessage, setSyncMessage] = useState('');

    // Multi-select Class Filter (Summary & Sort Ngang)
    const [selectedClasses, setSelectedClasses] = useState<string[]>([]);
    const [isFilterOpen, setIsFilterOpen] = useState(false);
    const filterRef = useRef<HTMLDivElement>(null);

    // Multi-select Student Filter (Summary & Sort Ngang)
    const [selectedStudentIds, setSelectedStudentIds] = useState<string[]>([]);
    const [isStudentFilterOpen, setIsStudentFilterOpen] = useState(false);
    const studentFilterRef = useRef<HTMLDivElement>(null);
    const [studentSearch, setStudentSearch] = useState('');

    // Click outside to close filters
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (filterRef.current && !filterRef.current.contains(event.target as Node)) {
                setIsFilterOpen(false);
            }
            if (studentFilterRef.current && !studentFilterRef.current.contains(event.target as Node)) {
                setIsStudentFilterOpen(false);
            }
            if (chartClassFilterRef.current && !chartClassFilterRef.current.contains(event.target as Node)) {
                setIsChartClassFilterOpen(false);
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

        const password = window.prompt("Vui lòng nhập mật khẩu để xác nhận đồng bộ:");
        if (password !== "nguyenduchien") {
            alert("Mật khẩu không đúng! Hủy đồng bộ.");
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

    // Collect all unique classes from the student list for the filter dropdown
    const uniqueClasses = useMemo(() => {
        const classes = new Set<string>();
        if (Array.isArray(students)) {
            students.forEach(s => s && s.class && classes.add(s.class));
        }
        return Array.from(classes).sort();
    }, [students]);

    useEffect(() => {
        if (uniqueClasses.length > 0 && selectedChartClasses.length === 0) {
            setSelectedChartClasses([uniqueClasses[0]]);
        }
    }, [uniqueClasses, selectedChartClasses]);

    const activeChartClasses = useMemo(() => {
        return selectedChartClasses.length > 0 ? selectedChartClasses : (uniqueClasses.length > 0 ? [uniqueClasses[0]] : []);
    }, [selectedChartClasses, uniqueClasses]);

    // --- CHART TAB CALCULATIONS & HELPERS ---

    const getStudentExamScore = (sId: string, examIndex: number, type: string) => {
        const record = examData[examIndex]?.[sId];
        if (!record) return undefined;
        const student = students.find(s => s.id === sId);
        if (!student) return undefined;
        const block = getBlockType(student.class);

        if (type === 'math') return record.math;
        if (type === 'phys') return record.phys;
        if (type === 'chem') return record.chem;
        if (type === 'bio') return record.bio;
        if (type === 'eng') return record.eng;
        if (type === 'A') {
            if (record.math !== undefined && record.phys !== undefined && record.chem !== undefined) {
                return record.math + record.phys + record.chem;
            }
        }
        if (type === 'B') {
            if (record.math !== undefined && record.chem !== undefined && record.bio !== undefined) {
                return record.math + record.chem + record.bio;
            }
        }
        if (type === 'A1') {
            if (record.math !== undefined && record.phys !== undefined && record.eng !== undefined) {
                return record.math + record.phys + record.eng;
            }
        }
        if (type === 'total') {
            if (block === 'A') {
                if (record.math !== undefined && record.phys !== undefined && record.chem !== undefined) 
                    return record.math + record.phys + record.chem;
            } else if (block === 'B') {
                if (record.math !== undefined && record.chem !== undefined && record.bio !== undefined)
                    return record.math + record.chem + record.bio;
            } else if (block === 'A1') {
                if (record.math !== undefined && record.phys !== undefined && record.eng !== undefined)
                    return record.math + record.phys + record.eng;
            }
        }
        return undefined;
    };

    const rankingMap = useMemo(() => {
        const map: Record<number, Record<string, { rank: number; total: number; score: number }>> = {};
        
        for (let i = 1; i <= 40; i++) {
            map[i] = {};
            
            const studentsWithScores = students.map(s => {
                const score = getStudentExamScore(s.id, i, rankScoreType);
                return { s, score };
            }).filter(item => item.score !== undefined) as { s: StudentProfile; score: number }[];

            if (studentsWithScores.length === 0) continue;

            if (comparisonMode === 'class') {
                const classes: Record<string, typeof studentsWithScores> = {};
                studentsWithScores.forEach(item => {
                    const c = item.s.class;
                    if (!classes[c]) classes[c] = [];
                    classes[c].push(item);
                });
                
                Object.keys(classes).forEach(cName => {
                    const group = classes[cName];
                    group.sort((a, b) => b.score - a.score);
                    group.forEach((item) => {
                        const rank = group.filter(x => x.score > item.score).length + 1;
                        map[i][item.s.id] = { rank, total: group.length, score: item.score };
                    });
                });
            } else if (comparisonMode === 'block') {
                const blocks: Record<string, typeof studentsWithScores> = {};
                studentsWithScores.forEach(item => {
                    const b = getBlockType(item.s.class);
                    if (!blocks[b]) blocks[b] = [];
                    blocks[b].push(item);
                });
                
                Object.keys(blocks).forEach(bName => {
                    const group = blocks[bName];
                    group.sort((a, b) => b.score - a.score);
                    group.forEach((item) => {
                        const rank = group.filter(x => x.score > item.score).length + 1;
                        map[i][item.s.id] = { rank, total: group.length, score: item.score };
                    });
                });
            } else {
                studentsWithScores.sort((a, b) => b.score - a.score);
                studentsWithScores.forEach((item) => {
                    const rank = studentsWithScores.filter(x => x.score > item.score).length + 1;
                    map[i][item.s.id] = { rank, total: studentsWithScores.length, score: item.score };
                });
            }
        }
        return map;
    }, [students, examData, rankScoreType, comparisonMode]);

    const chartClassStudents = useMemo(() => {
        return students.filter(s => s && activeChartClasses.includes(s.class));
    }, [students, activeChartClasses]);

    const activeExams = useMemo(() => {
        const exams: number[] = [];
        for (let i = 1; i <= 40; i++) {
            const hasScore = chartClassStudents.some(s => rankingMap[i]?.[s.id] !== undefined);
            if (hasScore) {
                exams.push(i);
            }
        }
        return exams.sort((a, b) => a - b);
    }, [chartClassStudents, rankingMap]);

    useEffect(() => {
        if (chartClassStudents.length > 0) {
            const studentsWithAverages = chartClassStudents.map(s => {
                const scs: number[] = [];
                for (let i = 1; i <= 40; i++) {
                    const score = getStudentExamScore(s.id, i, rankScoreType);
                    if (score !== undefined) scs.push(score);
                }
                const avg = scs.length > 0 ? (scs.reduce((a, b) => a + b, 0) / scs.length) : 0;
                return { id: s.id, avg };
            }).sort((a, b) => b.avg - a.avg);

            const initialVisible: Record<string, boolean> = {};
            studentsWithAverages.forEach((item, idx) => {
                initialVisible[item.id] = idx < 10;
            });
            setVisibleStudents(initialVisible);
        }
    }, [activeChartClasses, rankScoreType, students]);

    const colorsList = [
        '#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', 
        '#ec4899', '#14b8a6', '#f97316', '#06b6d4', '#6366f1',
        '#059669', '#d97706', '#dc2626', '#7c3aed', '#db2777', 
        '#0d9488', '#ea580c', '#0891b2', '#4f46e5'
    ];
    const getStudentColor = (idx: number) => colorsList[idx % colorsList.length];

    const getComputedData = useMemo(() => {
        if (!Array.isArray(students)) return [];
        let filteredStudents = selectedClasses.length > 0 ? students.filter(s => s && selectedClasses.includes(s.class)) : students;
        if (selectedStudentIds.length > 0) {
            filteredStudents = filteredStudents.filter(s => s && selectedStudentIds.includes(s.id));
        }
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
            else if (summaryTab === 'custom') shouldInclude = true;
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
            else if (summaryTab === 'custom') {
                if (customSubjects.math) finalVal += avgMath;
                if (customSubjects.phys) finalVal += avgPhys;
                if (customSubjects.chem) finalVal += avgChem;
                if (customSubjects.eng) finalVal += avgEng;
            }
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
                    } else if (summaryTab === 'custom') {
                         let sum = 0;
                         let hasValue = false;
                         if (customSubjects.math && record.math !== undefined) { sum += record.math; hasValue = true; }
                         if (customSubjects.phys && record.phys !== undefined) { sum += record.phys; hasValue = true; }
                         if (customSubjects.chem && record.chem !== undefined) { sum += record.chem; hasValue = true; }
                         if (customSubjects.eng && record.eng !== undefined) { sum += record.eng; hasValue = true; }
                         if (hasValue) colVal = sum;
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



    // --- FIX: Define activeExamScoreList ---
    const activeExamScoreList = useMemo(() => {
        let list = students || [];
        if (selectedClasses.length > 0) {
            list = list.filter(s => s && selectedClasses.includes(s.class));
        }
        if (selectedStudentIds.length > 0) {
            list = list.filter(s => s && selectedStudentIds.includes(s.id));
        }
        return list.map(s => {
             if (!s) return null;
             return {
                ...s,
                scores: examData[activeExamTime]?.[s.id] || {}
            };
        }).filter(s => s !== null);
    }, [students, selectedClasses, selectedStudentIds, examData, activeExamTime]);
    // ---------------------------------------

    const availableStudents = useMemo(() => {
        let list = students || [];
        if (selectedClasses.length > 0) {
            list = list.filter(s => s && selectedClasses.includes(s.class));
        }
        return list;
    }, [students, selectedClasses]);

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

    const renderStudentSelect = () => {
        const displayList = availableStudents.filter(s => 
            s.fullName.toLowerCase().includes(studentSearch.toLowerCase()) || 
            s.id.includes(studentSearch) ||
            s.class.toLowerCase().includes(studentSearch.toLowerCase())
        );

        return (
            <div style={{ position: 'relative' }} ref={studentFilterRef}>
                <button
                    onClick={() => setIsStudentFilterOpen(!isStudentFilterOpen)}
                    style={{
                        padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1', background: 'white',
                        fontSize: '13px', color: '#475569', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', minWidth: '160px'
                    }}
                >
                    <Users size={14} />
                    {selectedStudentIds.length === 0 ? "Tất cả học sinh" : `Đang chọn ${selectedStudentIds.length} học sinh`}
                </button>
                
                {isStudentFilterOpen && (
                    <div style={{
                        position: 'absolute', top: '100%', right: 0, marginTop: '5px', background: 'white',
                        border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                        zIndex: 100, width: '280px', maxHeight: '350px', display: 'flex', flexDirection: 'column', padding: '10px'
                    }}>
                        <input
                            type="text"
                            placeholder="Tìm tên, SBD, lớp..."
                            value={studentSearch}
                            onChange={(e) => setStudentSearch(e.target.value)}
                            style={{ padding: '6px 10px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '12px', marginBottom: '8px' }}
                        />
                        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px', fontSize: '12px' }}>
                            <span 
                                style={{ cursor: 'pointer', color: '#3b82f6', fontWeight: 600 }}
                                onClick={() => setSelectedStudentIds(availableStudents.map(s => s.id))}
                            >
                                Chọn tất cả
                            </span>
                            <span 
                                style={{ cursor: 'pointer', color: '#64748b' }}
                                onClick={() => setSelectedStudentIds([])}
                            >
                                Bỏ chọn
                            </span>
                        </div>
                        <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '2px', maxHeight: '220px' }}>
                            {displayList.map(s => (
                                <div key={s.id} style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '4px 0', fontSize: '12px' }}>
                                    <input
                                        type="checkbox"
                                        id={`filter-std-${s.id}`}
                                        checked={selectedStudentIds.includes(s.id)}
                                        onChange={(e) => {
                                            if (e.target.checked) setSelectedStudentIds(prev => [...prev, s.id]);
                                            else setSelectedStudentIds(prev => prev.filter(id => id !== s.id));
                                        }}
                                        style={{ cursor: 'pointer' }}
                                    />
                                    <label htmlFor={`filter-std-${s.id}`} style={{ cursor: 'pointer', flex: 1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                        <span style={{ fontWeight: 600 }}>{s.fullName}</span> <span style={{ color: '#64748b', fontSize: '11px' }}>({s.class})</span>
                                    </label>
                                </div>
                            ))}
                            {displayList.length === 0 && <div style={{ color: '#94a3b8', fontSize: '12px', fontStyle: 'italic', padding: '8px 0' }}>Không tìm thấy học sinh</div>}
                        </div>
                    </div>
                )}
            </div>
        );
    };

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
                    onClick={() => setSubTab('chart')}
                    style={{ 
                        padding: '12px', borderRadius: '8px', border: 'none', cursor: 'pointer', textAlign: 'left',
                        background: subTab === 'chart' ? '#eff6ff' : 'transparent',
                        color: subTab === 'chart' ? '#1e3a8a' : '#64748b', fontWeight: subTab === 'chart' ? 600 : 500,
                        display: 'flex', alignItems: 'center', gap: '10px'
                    }}
                >
                    <TrendingUp size={18} /> Biểu đồ
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


            </div>

            <div style={{ flex: 1, padding: '24px', overflow: 'hidden', display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                
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
                            <div style={{ flex: 1, overflow: 'auto', width: '100%' }}>
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

                         <div style={{ flex: 1, overflow: 'auto', background: '#f8fafc', padding: '24px', width: '100%' }}>
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
                                 {id: 'custom', label: 'Kết hợp'},
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
                             
                             {summaryTab === 'custom' && (
                                  <div style={{ display: 'flex', alignItems: 'center', gap: '15px', padding: '6px 12px', background: '#eff6ff', borderRadius: '8px', border: '1px solid #bfdbfe' }}>
                                      <span style={{ fontSize: '13px', fontWeight: 600, color: '#1e40af' }}>Môn kết hợp:</span>
                                      {[
                                          { id: 'math', label: 'Toán' },
                                          { id: 'phys', label: 'Lí' },
                                          { id: 'chem', label: 'Hóa' },
                                          { id: 'eng', label: 'Anh' }
                                      ].map(sub => (
                                          <label key={sub.id} style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: '#1e3a8a', cursor: 'pointer', fontWeight: 600 }}>
                                              <input 
                                                  type="checkbox" 
                                                  checked={customSubjects[sub.id as keyof typeof customSubjects]} 
                                                  onChange={(e) => setCustomSubjects(prev => ({ ...prev, [sub.id]: e.target.checked }))}
                                                  style={{ cursor: 'pointer', width: '15px', height: '15px' }}
                                              />
                                              {sub.label}
                                          </label>
                                      ))}
                                  </div>
                             )}
                             
                             <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px', alignItems: 'center' }}>
                                {renderMultiSelect()}
                                {renderStudentSelect()}
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

                        <div style={{ flex: 1, overflow: 'auto', width: '100%' }}>
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
                    <div style={{ height: '100%', display: 'flex', justifyContent: 'center', alignItems: 'center', padding: '20px' }}>
                        <div style={{ width: '100%', maxWidth: '600px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                            <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '24px' }}>
                                    <div style={{ width: '40px', height: '40px', borderRadius: '10px', background: '#ecfdf5', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#059669' }}>
                                        <Share2 size={24} />
                                    </div>
                                    <div>
                                        <h3 style={{ margin: 0, color: '#1e293b', fontSize: '18px', fontWeight: 600 }}>Cấu hình kết nối Google Sheets</h3>
                                        <p style={{ margin: '4px 0 0 0', color: '#64748b', fontSize: '13px' }}>Kết nối để đồng bộ dữ liệu điểm và học sinh lên đám mây.</p>
                                    </div>
                                </div>
                                
                                <div style={{ marginBottom: '24px' }}>
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
                                <h4 style={{ margin: '0 0 12px 0', color: '#475569', fontSize: '14px', fontWeight: 600 }}>Trạng thái dữ liệu hiện tại</h4>
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
                    </div>
                )}
                
                {subTab === 'chart' && (
                    <div style={{ display: 'flex', gap: '20px', height: '100%', padding: '10px', overflow: 'hidden', flex: 1 }}>
                        {/* Left Column: Chart Controls and SVG */}
                        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '16px', background: 'white', padding: '20px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', overflow: 'hidden' }}>
                            
                            {/* Controls row */}
                            <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: '16px' }}>
                                {/* Choose class (multi-select) */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Lớp học</span>
                                    <div style={{ position: 'relative' }} ref={chartClassFilterRef}>
                                        <button
                                            onClick={() => setIsChartClassFilterOpen(!isChartClassFilterOpen)}
                                            style={{
                                                padding: '8px 12px', borderRadius: '8px', border: '1px solid #cbd5e1', background: 'white',
                                                fontSize: '13px', color: '#334155', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', minWidth: '150px', fontWeight: 500
                                            }}
                                        >
                                            <Filter size={14} />
                                            {activeChartClasses.length === 0 ? "Chọn lớp" : activeChartClasses.length === 1 ? `Lớp ${activeChartClasses[0]}` : `Đang chọn ${activeChartClasses.length} lớp`}
                                        </button>
                                        {isChartClassFilterOpen && (
                                            <div style={{
                                                position: 'absolute', top: '100%', left: 0, marginTop: '5px', background: 'white',
                                                border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)',
                                                zIndex: 100, width: '220px', maxHeight: '280px', overflowY: 'auto', padding: '10px'
                                            }}>
                                                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px', fontSize: '12px' }}>
                                                    <span 
                                                        style={{ cursor: 'pointer', color: '#3b82f6', fontWeight: 600 }}
                                                        onClick={() => setSelectedChartClasses(uniqueClasses)}
                                                    >
                                                        Chọn tất cả
                                                    </span>
                                                    <span 
                                                        style={{ cursor: 'pointer', color: '#64748b' }}
                                                        onClick={() => setSelectedChartClasses([])}
                                                    >
                                                        Bỏ chọn
                                                    </span>
                                                </div>
                                                {uniqueClasses.map(cls => (
                                                    <div key={cls} style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '5px 0', fontSize: '13px' }}>
                                                        <input
                                                            type="checkbox"
                                                            id={`chart-cls-${cls}`}
                                                            checked={activeChartClasses.includes(cls)}
                                                            onChange={(e) => {
                                                                if (e.target.checked) setSelectedChartClasses(prev => [...prev, cls]);
                                                                else setSelectedChartClasses(prev => prev.filter(c => c !== cls));
                                                            }}
                                                            style={{ cursor: 'pointer' }}
                                                        />
                                                        <label htmlFor={`chart-cls-${cls}`} style={{ cursor: 'pointer', flex: 1 }}>{cls}</label>
                                                    </div>
                                                ))}
                                            </div>
                                        )}
                                    </div>
                                </div>

                                {/* Choose rank score type */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Điểm tính hạng</span>
                                    <select
                                        value={rankScoreType}
                                        onChange={(e) => setRankScoreType(e.target.value as any)}
                                        style={{ padding: '8px 12px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '13px', background: 'white', minWidth: '180px', cursor: 'pointer', fontWeight: 500, color: '#334155' }}
                                    >
                                        <option value="total">Tổng Khối (Tự động)</option>
                                        <option value="math">Toán</option>
                                        <option value="phys">Lí</option>
                                        <option value="chem">Hóa</option>
                                        <option value="eng">Anh</option>
                                        <option value="bio">Sinh</option>
                                        <option value="A">Khối A (T-L-H)</option>
                                        <option value="A1">Khối A1 (T-L-A)</option>
                                        <option value="B">Khối B (T-H-S)</option>
                                    </select>
                                </div>

                                {/* Choose comparison group */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Quy chiếu xếp hạng</span>
                                    <div style={{ display: 'flex', border: '1px solid #cbd5e1', borderRadius: '8px', overflow: 'hidden' }}>
                                        {[
                                            { id: 'class', label: 'Riêng trong Lớp' },
                                            { id: 'block', label: 'Trong Khối' },
                                            { id: 'school', label: 'Toàn Trường' }
                                        ].map(mode => (
                                            <button
                                                key={mode.id}
                                                onClick={() => setComparisonMode(mode.id as any)}
                                                style={{
                                                    padding: '8px 16px', border: 'none', fontSize: '13px', cursor: 'pointer', fontWeight: 600,
                                                    background: comparisonMode === mode.id ? '#1e3a8a' : 'white',
                                                    color: comparisonMode === mode.id ? 'white' : '#475569',
                                                    transition: 'all 0.15s ease'
                                                }}
                                            >
                                                {mode.label}
                                            </button>
                                        ))}
                                    </div>
                                </div>

                                {/* Full scale toggle */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Trục đứng (Thứ hạng)</span>
                                    <label style={{ 
                                        display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 12px', 
                                        borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '13px', 
                                        background: 'white', cursor: 'pointer', fontWeight: 500, color: '#334155',
                                        minHeight: '38px', userSelect: 'none'
                                    }}>
                                        <input
                                            type="checkbox"
                                            checked={useFullScale}
                                            onChange={(e) => setUseFullScale(e.target.checked)}
                                            style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                                        />
                                        Hiển thị hết sỉ số ({
                                            comparisonMode === 'class' 
                                                ? students.filter(s => s && activeChartClasses.includes(s.class)).length 
                                                : comparisonMode === 'block'
                                                    ? students.filter(s => s && activeChartClasses.map(c => getBlockType(c)).includes(getBlockType(s.class))).length
                                                    : students.length
                                        })
                                    </label>
                                </div>

                                {/* Show all labels toggle */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Số hạng trên đường</span>
                                    <label style={{ 
                                        display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 12px', 
                                        borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '13px', 
                                        background: 'white', cursor: 'pointer', fontWeight: 500, color: '#334155',
                                        minHeight: '38px', userSelect: 'none'
                                    }}>
                                        <input
                                            type="checkbox"
                                            checked={showAllLabels}
                                            onChange={(e) => setShowAllLabels(e.target.checked)}
                                            style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                                        />
                                        Hiện số hạng
                                    </label>
                                </div>

                                {/* Chart Height Slider */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', minWidth: '180px' }}>
                                    <span style={{ fontSize: '12px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Chiều cao biểu đồ: {chartHeight}px</span>
                                    <div style={{ 
                                        display: 'flex', alignItems: 'center', gap: '8px', padding: '8px 12px', 
                                        borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '13px', 
                                        background: 'white', minHeight: '38px', userSelect: 'none'
                                    }}>
                                        <input
                                            type="range"
                                            min="250"
                                            max="850"
                                            step="10"
                                            value={chartHeight}
                                            onChange={(e) => setChartHeight(Number(e.target.value))}
                                            style={{ cursor: 'pointer', flex: 1 }}
                                        />
                                    </div>
                                </div>
                            </div>

                            {/* Chart Area */}
                            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', position: 'relative', overflow: 'hidden' }}>
                                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                                    <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 600, color: '#1e293b' }}>
                                        Biểu đồ quá trình học tập lớp {activeChartClasses.join(', ')} ({activeExams.length} lần thi)
                                    </h3>
                                    <span style={{ fontSize: '12px', color: '#64748b', fontStyle: 'italic' }}>
                                        * Trục đứng biểu diễn thứ hạng (hạng nhỏ ở trên là cao hơn)
                                    </span>
                                </div>
                                <div style={{ height: `${chartHeight}px`, position: 'relative', display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', overflow: 'hidden' }}>
                                    {activeExams.length === 0 ? (
                                        <div style={{ textAlign: 'center', color: '#94a3b8' }}>
                                            <TrendingUp size={48} style={{ margin: '0 auto 12px', opacity: 0.3 }} />
                                            <p style={{ margin: 0, fontWeight: 600 }}>Chưa có dữ liệu điểm để vẽ biểu đồ cho lớp này.</p>
                                            <p style={{ margin: '4px 0 0 0', fontSize: '12px', color: '#cbd5e1' }}>Vui lòng kiểm tra dữ liệu điểm trong tab "Dữ liệu điểm".</p>
                                        </div>
                                    ) : (
                                        <ResponsiveSVGChart 
                                            students={students}
                                            chartClassStudents={chartClassStudents}
                                            activeExams={activeExams}
                                            rankingMap={rankingMap}
                                            visibleStudents={visibleStudents}
                                            hoveredStudentId={hoveredStudentId}
                                            setHoveredStudentId={setHoveredStudentId}
                                            useFullScale={useFullScale}
                                            comparisonMode={comparisonMode}
                                            selectedChartClasses={activeChartClasses}
                                            getStudentColor={getStudentColor}
                                            showAllLabels={showAllLabels}
                                        />
                                    )}
                                </div>
                            </div>
                        </div>

                        {/* Right Column: Student List with toggleable checkbox */}
                        <div style={{ width: '320px', display: 'flex', flexDirection: 'column', background: 'white', borderRadius: '12px', border: '1px solid #e2e8f0', overflow: 'hidden' }}>
                            <div style={{ padding: '16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0' }}>
                                <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', color: '#334155', fontWeight: 700 }}>
                                    Học sinh ({chartClassStudents.length})
                                </h4>
                                
                                {/* Search filter */}
                                <input 
                                    type="text" 
                                    placeholder="Tìm tên học sinh..."
                                    value={chartSearch}
                                    onChange={(e) => setChartSearch(e.target.value)}
                                    style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '13px', marginBottom: '12px' }}
                                />

                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    <div style={{ display: 'flex', gap: '6px' }}>
                                        <button
                                            onClick={() => {
                                                const newVisible: Record<string, boolean> = {};
                                                chartClassStudents.forEach(s => newVisible[s.id] = true);
                                                setVisibleStudents(newVisible);
                                            }}
                                            style={{ flex: 1, padding: '6px 8px', background: 'white', border: '1px solid #cbd5e1', borderRadius: '6px', fontSize: '11px', fontWeight: 600, color: '#475569', cursor: 'pointer' }}
                                        >
                                            Hiện hết
                                        </button>
                                        <button
                                            onClick={() => {
                                                setVisibleStudents({});
                                            }}
                                            style={{ flex: 1, padding: '6px 8px', background: 'white', border: '1px solid #cbd5e1', borderRadius: '6px', fontSize: '11px', fontWeight: 600, color: '#475569', cursor: 'pointer' }}
                                        >
                                            Ẩn hết
                                        </button>
                                    </div>
                                    <div style={{ display: 'flex', gap: '6px' }}>
                                        <button
                                            onClick={() => {
                                                const studentsWithAverages = chartClassStudents.map(s => {
                                                    const scs: number[] = [];
                                                    for (let i = 1; i <= 40; i++) {
                                                        const score = getStudentExamScore(s.id, i, rankScoreType);
                                                        if (score !== undefined) scs.push(score);
                                                    }
                                                    const avg = scs.length > 0 ? (scs.reduce((a, b) => a + b, 0) / scs.length) : 0;
                                                    return { id: s.id, avg };
                                                }).sort((a, b) => b.avg - a.avg);

                                                const newVisible: Record<string, boolean> = {};
                                                studentsWithAverages.forEach((item, idx) => {
                                                    newVisible[item.id] = idx < 10;
                                                });
                                                setVisibleStudents(newVisible);
                                            }}
                                            style={{ flex: 1, padding: '6px 8px', background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '6px', fontSize: '11px', fontWeight: 700, color: '#1e40af', cursor: 'pointer' }}
                                        >
                                            Top 10 trên
                                        </button>
                                        <button
                                            onClick={() => {
                                                const studentsWithAverages = chartClassStudents.map(s => {
                                                    const scs: number[] = [];
                                                    for (let i = 1; i <= 40; i++) {
                                                        const score = getStudentExamScore(s.id, i, rankScoreType);
                                                        if (score !== undefined) scs.push(score);
                                                    }
                                                    const avg = scs.length > 0 ? (scs.reduce((a, b) => a + b, 0) / scs.length) : 0;
                                                    return { id: s.id, avg };
                                                }).sort((a, b) => a.avg - b.avg);

                                                const newVisible: Record<string, boolean> = {};
                                                studentsWithAverages.forEach((item, idx) => {
                                                    newVisible[item.id] = idx < 10;
                                                });
                                                setVisibleStudents(newVisible);
                                            }}
                                            style={{ flex: 1, padding: '6px 8px', background: '#fef2f2', border: '1px solid #fecaca', borderRadius: '6px', fontSize: '11px', fontWeight: 700, color: '#991b1b', cursor: 'pointer' }}
                                        >
                                            Top 10 dưới
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <div style={{ flex: 1, overflowY: 'auto', padding: '10px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                {chartClassStudents
                                    .filter(s => s.fullName.toLowerCase().includes(chartSearch.toLowerCase()) || s.id.includes(chartSearch))
                                    .map((s) => {
                                        const idxInClass = chartClassStudents.findIndex(x => x.id === s.id);
                                        const color = getStudentColor(idxInClass);
                                        const isChecked = !!visibleStudents[s.id];
                                        const isHovered = hoveredStudentId === s.id;
                                        
                                        const studentRanks: number[] = [];
                                        activeExams.forEach(ex => {
                                            const r = rankingMap[ex]?.[s.id]?.rank;
                                            if (r !== undefined) studentRanks.push(r);
                                        });
                                        const avgRank = studentRanks.length > 0 ? (studentRanks.reduce((a, b) => a + b, 0) / studentRanks.length).toFixed(1) : '-';

                                        return (
                                            <div
                                                key={s.id}
                                                onMouseEnter={() => setHoveredStudentId(s.id)}
                                                onMouseLeave={() => setHoveredStudentId(null)}
                                                style={{
                                                    display: 'flex', alignItems: 'center', gap: '10px', padding: '8px 12px', borderRadius: '8px',
                                                    background: isHovered ? '#f1f5f9' : 'transparent',
                                                    cursor: 'pointer', transition: 'background 0.1s ease',
                                                    border: isHovered ? '1px solid #e2e8f0' : '1px solid transparent'
                                                }}
                                            >
                                                <input
                                                    type="checkbox"
                                                    checked={isChecked}
                                                    onChange={(e) => {
                                                        setVisibleStudents(prev => ({ ...prev, [s.id]: e.target.checked }));
                                                    }}
                                                    style={{ cursor: 'pointer', width: '16px', height: '16px' }}
                                                />
                                                <div 
                                                    style={{ width: '12px', height: '12px', borderRadius: '50%', background: color, flexShrink: 0 }}
                                                    onClick={() => {
                                                        setVisibleStudents(prev => ({ ...prev, [s.id]: !isChecked }));
                                                    }}
                                                />
                                                <div 
                                                    style={{ flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column', gap: '2px' }}
                                                    onClick={() => {
                                                        setVisibleStudents(prev => ({ ...prev, [s.id]: !isChecked }));
                                                    }}
                                                >
                                                    <span style={{ fontSize: '13px', fontWeight: isChecked ? 600 : 500, color: '#334155', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                                        {s.fullName}
                                                        {activeChartClasses.length > 1 && (
                                                            <span style={{ fontSize: '10px', background: '#e2e8f0', color: '#475569', padding: '1px 5px', borderRadius: '4px', fontWeight: 600 }}>
                                                                {s.class}
                                                            </span>
                                                        )}
                                                    </span>
                                                    <span style={{ fontSize: '11px', color: '#94a3b8' }}>
                                                        SBD: {s.id}
                                                    </span>
                                                </div>
                                                <span style={{ fontSize: '11px', color: '#475569', background: '#f8fafc', padding: '3px 8px', borderRadius: '4px', border: '1px solid #f1f5f9', fontWeight: 600 }}>
                                                    Hạng: {avgRank}
                                                </span>
                                            </div>
                                        );
                                    })}
                            </div>
                        </div>
                    </div>
                )}
                


            </div>
        </div>
    );
};

const App = RankingView;

const root = createRoot(document.getElementById('root')!);
root.render(<App />);