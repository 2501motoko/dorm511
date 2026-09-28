"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  BookOpen, CalendarDays, Check, ChevronDown, ChevronLeft, ChevronRight,
  ClipboardCheck, Clock3, Download, Dumbbell, FileDown, FileText, Images,
  LayoutDashboard, LockKeyhole, MapPin, Pencil, Plus, Search, ShieldCheck, Trash2,
  Upload, Users, X,
} from "lucide-react";

type Member = { id: string; displayName: string; role: "admin" | "member" };
type Course = { id: string; name: string };
type CalendarItem = { id: string; kind: "ddl" | "activity"; title: string; courseId: string | null; course: string | null; date: string; dueTime: string | null; startTime: string | null; endTime: string | null; location: string | null; notes: string | null; isPublic: number; ownerId: string; owner: string };
type Resource = { id: string; courseId: string; course: string; title: string; fileName: string; contentType: string; size: number; uploaderId: string; uploader: string; createdAt: string };
type Duty = { id: string; date: string; garbageMemberId: string | null; sweepMemberId: string | null; garbageDone: number; sweepDone: number; garbageDoneBy: string | null; sweepDoneBy: string | null; garbageName: string | null; sweepName: string | null };
type Sport = { id: string; activityType: string | null; durationMinutes: number | null; date: string; isPublic: number; ownerId: string; owner: string };
type Memory = { id: string; caption: string | null; date: string | null; fileName: string; contentType: string; size: number; uploaderId: string; uploader: string; createdAt: string };
type Data = { member: Member; courses: Course[]; resources: Resource[]; calendar: CalendarItem[]; duties: Duty[]; sports: Sport[]; memories: Memory[]; members: Member[] };
type EditState = { type: "calendar"; item: CalendarItem } | { type: "resource"; item: Resource } | { type: "duty"; item: Duty } | { type: "sport"; item: Sport } | { type: "memory"; item: Memory };

const nav = [
  { id: "home", label: "总览", icon: LayoutDashboard },
  { id: "resources", label: "课程资料", icon: BookOpen },
  { id: "calendar", label: "DDL 与活动", icon: CalendarDays },
  { id: "duties", label: "值日表", icon: ClipboardCheck },
  { id: "sports", label: "运动打卡", icon: Dumbbell },
  { id: "memories", label: "宿舍风采", icon: Images },
] as const;
type Tab = (typeof nav)[number]["id"];

const shanghaiDate = () => new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai" }).format(new Date());
const humanDate = (value: string) => {
  if (!value) return "";
  const [y, m, d] = value.split("-").map(Number);
  return `${m}月${d}日`;
};
const prettySize = (n: number) => n >= 1024 * 1024 ? `${(n / 1024 / 1024).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`;

async function callApi(body?: Record<string, unknown>) {
  const response = await fetch("/api/data", body ? { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) } : {});
  const payload = await response.json() as { error?: string };
  if (!response.ok) throw new Error(payload.error || "操作没有完成，请重试。");
  return payload;
}

function esc(value: unknown) {
  return String(value ?? "").replace(/[&<>\"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" })[c]!);
}

// A small uncompressed ZIP writer creates a genuine .xlsx without an extra runtime dependency.
function makeXlsx(rows: string[][]) {
  const xmlRows = rows.map((row, r) => `<row r="${r + 1}">${row.map((cell, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}" t="inlineStr"><is><t xml:space="preserve">${esc(cell)}</t></is></c>`).join("")}</row>`).join("");
  const files: [string, string][] = [
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="DDL与活动" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`],
    ["xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${xmlRows}</sheetData></worksheet>`],
  ];
  const encoder = new TextEncoder();
  const crc32 = (bytes: Uint8Array) => { let crc = -1; for (const byte of bytes) { crc ^= byte; for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0); } return (crc ^ -1) >>> 0; };
  const local: Uint8Array[] = [], central: Uint8Array[] = []; let offset = 0;
  const u16 = (v: number) => [v & 255, (v >>> 8) & 255];
  const u32 = (v: number) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
  for (const [name, content] of files) {
    const n = encoder.encode(name), d = encoder.encode(content), crc = crc32(d);
    const head = new Uint8Array([0x50,0x4b,0x03,0x04,...u16(20),...u16(0x800),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(d.length),...u32(d.length),...u16(n.length),...u16(0),...n]);
    local.push(head, d);
    const dir = new Uint8Array([0x50,0x4b,0x01,0x02,...u16(20),...u16(20),...u16(0x800),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(d.length),...u32(d.length),...u16(n.length),...u16(0),...u16(0),...u16(0),...u16(0),...u32(0),...u32(offset),...n]);
    central.push(dir); offset += head.length + d.length;
  }
  const centralSize = central.reduce((sum, b) => sum + b.length, 0);
  const end = new Uint8Array([0x50,0x4b,0x05,0x06,...u16(0),...u16(0),...u16(files.length),...u16(files.length),...u32(centralSize),...u32(offset),...u16(0)]);
  const parts = [...local, ...central, end];
  const output = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let cursor = 0;
  for (const part of parts) { output.set(part, cursor); cursor += part.length; }
  return new Blob([output.buffer as ArrayBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function downloadCalendar(items: CalendarItem[]) {
  const rows = [["类型", "事项", "课程", "日期", "时间", "地点", "可见性", "备注", "创建者"], ...items.map((item) => [item.kind === "ddl" ? "DDL" : "活动", item.title, item.course ?? "", item.date, item.kind === "ddl" ? item.dueTime ?? "仅日期" : [item.startTime, item.endTime].filter(Boolean).join("–"), item.location ?? "", item.isPublic ? "全宿舍可见" : "仅自己可见", item.notes ?? "", item.owner])];
  const a = document.createElement("a"); a.href = URL.createObjectURL(makeXlsx(rows)); a.download = `511-日历-${shanghaiDate()}.xlsx`; a.click(); URL.revokeObjectURL(a.href);
}

function addToDeviceCalendar(item: CalendarItem) {
  const dateOnly = item.kind === "ddl" ? !item.dueTime : !item.startTime;
  const stamp = (date: string, time: string | null) => `${date.replaceAll("-", "")}T${(time ?? "00:00").replace(":", "")}00`;
  const start = item.kind === "ddl" ? stamp(item.date, item.dueTime) : stamp(item.date, item.startTime);
  const end = item.kind === "activity" && item.endTime ? stamp(item.date, item.endTime) : "";
  const escapeIcs = (value: string) => value.replaceAll("\\", "\\\\").replaceAll("\n", "\\n").replaceAll(",", "\\,").replaceAll(";", "\\;");
  const startLine = dateOnly ? `DTSTART;VALUE=DATE:${item.date.replaceAll("-", "")}` : `DTSTART;TZID=Asia/Shanghai:${start}`;
  const endLine = dateOnly ? `DTEND;VALUE=DATE:${nextDate(item.date)}` : end ? `DTEND;TZID=Asia/Shanghai:${end}` : "";
  const content = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//511 Dorm//ZH", "BEGIN:VEVENT", `UID:${item.id}@511-dorm`, `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, "")}`, startLine, endLine, `SUMMARY:${escapeIcs(item.title)}`, item.location ? `LOCATION:${escapeIcs(item.location)}` : "", item.notes ? `DESCRIPTION:${escapeIcs(item.notes)}` : "", "END:VEVENT", "END:VCALENDAR"].filter(Boolean).join("\r\n");
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([content], { type: "text/calendar;charset=utf-8" })); a.download = `${item.title.replace(/[^\p{L}\p{N}-]+/gu, "-")}.ics`; a.click(); URL.revokeObjectURL(a.href);
}

function nextDate(value: string) { const [y, m, d] = value.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10).replaceAll("-", ""); }

function IconButton({ label, onClick, children, danger = false }: { label: string; onClick: () => void; children: React.ReactNode; danger?: boolean }) {
  return <button title={label} aria-label={label} onClick={onClick} className={`icon-button ${danger ? "icon-danger" : ""}`}>{children}</button>;
}

export default function Home() {
  const [tab, setTab] = useState<Tab>("home");
  const [data, setData] = useState<Data | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [courseFilter, setCourseFilter] = useState<string[]>([]);
  const [month, setMonth] = useState(() => shanghaiDate().slice(0, 7));
  const [selectedDate, setSelectedDate] = useState(shanghaiDate());
  const [resourceSearch, setResourceSearch] = useState("");
  const [resourceCourse, setResourceCourse] = useState("all");
  const [showForm, setShowForm] = useState("");
  const [editing, setEditing] = useState<EditState | null>(null);

  const refresh = useCallback(async () => {
    setError("");
    try { const result = await callApi(); setData(result as Data); }
    catch (err) { setError(err instanceof Error ? err.message : "暂时无法连接共享数据。"); }
  }, []);
  useEffect(() => { void refresh(); }, [refresh]);
  useEffect(() => { if (!notice) return; const timer = setTimeout(() => setNotice(""), 3200); return () => clearTimeout(timer); }, [notice]);

  const mutate = async (body: Record<string, unknown>, message = "已保存"): Promise<boolean> => {
    setBusy(true); setError("");
    try { await callApi(body); await refresh(); setNotice(message); setShowForm(""); setEditing(null); return true; }
    catch (err) { setError(err instanceof Error ? err.message : "保存失败，请重试。"); return false; }
    finally { setBusy(false); }
  };
  const upload = async (form: FormData, message: string) => {
    setBusy(true); setError("");
    try {
      const response = await fetch("/api/files", { method: "POST", body: form }); const payload = await response.json() as { error?: string };
      if (!response.ok) throw new Error(payload.error || "上传失败。");
      await refresh(); setNotice(message); setShowForm("");
    } catch (err) { setError(err instanceof Error ? err.message : "上传失败，请重试。"); }
    finally { setBusy(false); }
  };

  const items = data?.calendar ?? [];
  const filteredCalendar = useMemo(() => items.filter((item) => item.kind !== "ddl" || courseFilter.length === 0 || courseFilter.includes(item.courseId ?? "")), [items, courseFilter]);
  const monthItems = filteredCalendar.filter((item) => item.date.startsWith(month));
  const selectedItems = filteredCalendar.filter((item) => item.date === selectedDate);
  const upcoming = [...items].filter((item) => item.date >= shanghaiDate()).sort((a, b) => `${a.date}${a.dueTime ?? a.startTime ?? ""}`.localeCompare(`${b.date}${b.dueTime ?? b.startTime ?? ""}`)).slice(0, 5);
  const dutyToday = data?.duties.find((d) => d.date === shanghaiDate());
  const doneCount = (data?.duties ?? []).reduce((sum, d) => sum + Number(!!d.garbageDone) + Number(!!d.sweepDone), 0);
  const totalChecks = (data?.duties ?? []).length * 2;

  const remove = async (action: string, id: string, confirmText: string) => {
    if (!window.confirm(confirmText)) return;
    await mutate({ action, id }, "已删除");
  };
  const saveEdit = (fields: Record<string, unknown>) => {
    if (!editing) return;
    const action: Record<EditState["type"], string> = { calendar: "editCalendar", resource: "editResource", duty: "editDuty", sport: "editSport", memory: "editMemory" };
    void mutate({ action: action[editing.type], id: editing.item.id, ...fields }, "修改已保存");
  };
  const title = nav.find((item) => item.id === tab)?.label ?? "总览";

  if (error && !data) return <main className="auth-screen"><section className="auth-card"><div className="brand-mark">511</div><p className="eyebrow">511 SOCIOLOGY DORM</p><h1>进入宿舍共享空间</h1><p className="muted">登录后才能查看课程资料、日历和宿舍记录。</p><a className="primary-button sign-in" href="/signin-with-chatgpt?return_to=%2F" target="_top">登录后继续</a><p className="auth-error">{error}</p></section></main>;

  return <div className="app-shell">
    <aside className="sidebar">
      <a className="brand" href="#home" onClick={(e) => { e.preventDefault(); setTab("home"); }}><span className="brand-mark">511</span><span><b>社会学宿舍</b><small>SHARED SPACE</small></span></a>
      <div className="side-label">宿舍空间</div>
      <nav className="side-nav" aria-label="主导航">{nav.map(({ id, label, icon: Icon }) => <button key={id} className={`nav-item ${tab === id ? "active" : ""}`} onClick={() => { setTab(id); setShowForm(""); }}><Icon size={18} strokeWidth={1.9}/><span>{label}</span>{id === "calendar" && items.length > 0 && <i>{items.filter((item) => item.date >= shanghaiDate()).length}</i>}</button>)}</nav>
      <div className="sidebar-bottom"><div className="member-badge"><span className="avatar">{data?.member.displayName?.slice(0, 1) ?? "?"}</span><span><b>{data?.member.displayName ?? "正在连接…"}</b><small>{data?.member.role === "admin" ? "宿舍管理员" : "宿舍成员"}</small></span>{data?.member.role === "admin" && <ShieldCheck size={16} className="admin-icon"/>}</div><p>仅宿舍成员可访问</p></div>
    </aside>

    <main className="main-area">
      <header className="topbar"><div><div className="crumb">511 宿舍 <span>/</span> {title}</div><h1>{title}</h1></div><div className="top-actions"><div className="today-chip"><span className="today-dot"/>{humanDate(shanghaiDate())} <small>中国标准时间</small></div><button className="top-add" onClick={() => { setTab("calendar"); setShowForm("ddl"); }}><Plus size={17}/> 新增事项</button></div></header>
      <div className="mobile-nav">{nav.map(({ id, label, icon: Icon }) => <button key={id} className={tab === id ? "active" : ""} onClick={() => { setTab(id); setShowForm(""); }}><Icon size={16}/>{label}</button>)}</div>
      {error && <div className="error-banner" role="alert">{error}<button onClick={() => setError("")}>关闭</button></div>}
      {!data ? <section className="loading-card"><span className="spinner"/><p>正在加载宿舍共享数据…</p></section> : <>
        {tab === "home" && <section className="page-content">
          <div className="welcome-row"><div><p className="eyebrow">宿舍共享空间 · {new Date().getFullYear()}</p><h2>把共同生活，<em>记在一起。</em></h2><p>课程资料、待办日历和值日安排都在这里。</p></div><div className="week-stamp"><CalendarDays size={18}/><span>{humanDate(shanghaiDate())}</span><small>今天</small></div></div>
          <div className="metric-grid"><button onClick={() => setTab("calendar")} className="metric-card"><span className="metric-icon coral"><CalendarDays size={18}/></span><span className="metric-label">接下来的 DDL / 活动</span><b>{items.filter((item) => item.date >= shanghaiDate()).length}<small> 项</small></b><span className="metric-foot">查看共享日历 <ChevronRight size={14}/></span></button><button onClick={() => setTab("resources")} className="metric-card"><span className="metric-icon blue"><BookOpen size={18}/></span><span className="metric-label">课程资料</span><b>{data.resources.length}<small> 份</small></b><span className="metric-foot">按课程浏览 <ChevronRight size={14}/></span></button><button onClick={() => setTab("duties")} className="metric-card"><span className="metric-icon green"><ClipboardCheck size={18}/></span><span className="metric-label">值日完成情况</span><b>{doneCount}<small> / {totalChecks} 项</small></b><span className="metric-foot">倒垃圾与扫地分别打卡 <ChevronRight size={14}/></span></button></div>
          <div className="home-grid"><section className="panel upcoming-panel"><div className="panel-head"><div><p className="eyebrow">UP NEXT</p><h3>最近的日程</h3></div><button className="text-button" onClick={() => setTab("calendar")}>打开日历 <ChevronRight size={15}/></button></div>{upcoming.length ? <div className="upcoming-list">{upcoming.map((item) => <CalendarRow key={item.id} item={item} memberId={data.member.id} admin={data.member.role === "admin"} onEdit={() => setEditing({ type: "calendar", item })} onDelete={() => remove("deleteCalendar", item.id, "删除这条日程？")} onCalendar={() => addToDeviceCalendar(item)}/>)}</div> : <Empty title="暂时没有待办" text="新增课程 DDL 或活动，宿舍日历就会显示在这里。" action="新增 DDL" onClick={() => setShowForm("ddl")} />}</section>
            <section className="panel today-panel"><div className="panel-head"><div><p className="eyebrow">TODAY'S DUTY</p><h3>今日值日</h3></div><button className="soft-tag" onClick={() => setTab("duties")}>查看排班 <ChevronRight size={13}/></button></div>{dutyToday ? <div className="today-duty"><DutyTask label="倒垃圾" name={dutyToday.garbageName} done={!!dutyToday.garbageDone} me={dutyToday.garbageMemberId === data.member.id} admin={data.member.role === "admin"} onToggle={(done) => mutate({ action: "toggleDuty", id: dutyToday.id, field: "garbage", done }, done ? "已完成倒垃圾打卡" : "已撤销打卡")}/><DutyTask label="扫地" name={dutyToday.sweepName} done={!!dutyToday.sweepDone} me={dutyToday.sweepMemberId === data.member.id} admin={data.member.role === "admin"} onToggle={(done) => mutate({ action: "toggleDuty", id: dutyToday.id, field: "sweep", done }, done ? "已完成扫地打卡" : "已撤销打卡")}/><div className={`all-done ${dutyToday.garbageDone && dutyToday.sweepDone ? "complete" : ""}`}><span>{dutyToday.garbageDone && dutyToday.sweepDone ? "✓" : "·"}</span>{dutyToday.garbageDone && dutyToday.sweepDone ? "今日值日全部完成" : "两项都完成后，今日值日打卡成功"}</div></div> : <Empty title="今天没有排班" text="管理员可在值日表中安排倒垃圾和扫地。" action="查看值日表" onClick={() => setTab("duties")} />}</section></div>
        </section>}

        {tab === "resources" && <section className="page-content">
          <div className="section-intro"><div><p className="eyebrow">SHARED COURSE FILES</p><h2>课程资料</h2><p>按专业课整理，宿舍成员都可以上传和下载。</p></div><button className="primary-button" onClick={() => setShowForm(showForm === "resource" ? "" : "resource")}><Upload size={16}/>上传资料</button></div>
          {data.member.role === "admin" && <section className="admin-strip"><span className="admin-strip-icon"><ShieldCheck size={16}/></span><div><b>课程标签管理</b><small>只有管理员可以新建或删除课程标签</small></div><form className="inline-course-form" onSubmit={(e) => { e.preventDefault(); const form = e.currentTarget; const input = new FormData(form).get("course")?.toString().trim(); if (input) void mutate({ action: "addCourse", name: input }, "课程标签已创建").then((ok) => { if (ok) form.reset(); }); }}><input aria-label="课程名称" name="course" placeholder="输入课程名称" maxLength={80}/><button disabled={busy}><Plus size={15}/>添加标签</button></form>{data.courses.length > 0 && <div className="admin-course-list">{data.courses.map((course) => <span className="admin-course-tag" key={course.id}>{course.name}<IconButton label={`删除课程标签 ${course.name}`} danger onClick={() => remove("deleteCourse", course.id, `删除课程标签“${course.name}”？该课程已有资料或 DDL 时无法删除。`)}><X size={13}/></IconButton></span>)}</div>}</section>}
          {showForm === "resource" && <section className="form-panel"><div className="form-title"><div><p className="eyebrow">UPLOAD</p><h3>上传课程资料</h3></div><IconButton label="关闭" onClick={() => setShowForm("")}><X size={18}/></IconButton></div><form onSubmit={(e) => { e.preventDefault(); const form = new FormData(e.currentTarget); form.set("type", "resource"); void upload(form, "课程资料已上传"); }} className="form-grid"><label>课程标签<select name="courseId" required defaultValue=""><option value="" disabled>选择课程</option>{data.courses.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></label><label>资料名称<input name="title" placeholder="例如：第三周课堂讲义" maxLength={160} required/></label><label className="full-field file-drop"><Upload size={18}/><span>选择文件</span><small>PDF、Office、CSV、TXT、Markdown 或 ZIP · 单个文件不超过 25 MB</small><input name="file" type="file" accept=".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.csv,.txt,.md,.zip" required/></label><div className="form-actions"><button className="primary-button" disabled={busy}><Upload size={15}/>{busy ? "上传中…" : "上传到共享资料夹"}</button><span>其他成员可查看并下载</span></div></form></section>}
          {!data.courses.length && <div className="empty-course"><BookOpen size={24}/><h3>还没有课程标签</h3><p>管理员先添加公共专业课，成员就可以把资料归档到对应课程。</p></div>}
          <div className="resource-toolbar"><div className="search-box"><Search size={16}/><input value={resourceSearch} onChange={(e) => setResourceSearch(e.target.value)} placeholder="搜索资料名称或文件名"/></div><select value={resourceCourse} onChange={(e) => setResourceCourse(e.target.value)} aria-label="按课程筛选"><option value="all">全部课程</option>{data.courses.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select><span>{data.resources.length} 份资料</span></div>
          <div className="resource-grid">{data.resources.filter((r) => (resourceCourse === "all" || r.courseId === resourceCourse) && `${r.title} ${r.fileName} ${r.course}`.toLowerCase().includes(resourceSearch.toLowerCase())).map((r) => <article key={r.id} className="resource-card"><div className="file-icon"><FileText size={19}/></div><div className="resource-card-main"><span className="course-pill">{r.course}</span><h3>{r.title}</h3><p>{r.fileName}</p><small>{prettySize(r.size)} · {r.uploader} · {humanDate(r.createdAt.slice(0, 10))}</small></div><div className="resource-actions"><a className="download-button" href={`/api/files?id=${encodeURIComponent(r.id)}`}><Download size={16}/><span>下载</span></a>{(r.uploaderId === data.member.id || data.member.role === "admin") && <IconButton label="编辑资料" onClick={() => setEditing({ type: "resource", item: r })}><Pencil size={15}/></IconButton>}{(r.uploaderId === data.member.id || data.member.role === "admin") && <IconButton label="删除资料" danger onClick={() => remove("deleteResource", r.id, "删除这份共享资料？") }><Trash2 size={16}/></IconButton>}</div></article>)}</div>
          {!data.resources.length && data.courses.length > 0 && <Empty title="资料夹还是空的" text="上传课堂讲义、阅读材料或课程笔记，大家都能下载。" action="上传第一份资料" onClick={() => setShowForm("resource")} />}
        </section>}

        {tab === "calendar" && <section className="page-content">
          <div className="section-intro"><div><p className="eyebrow">DEADLINES & DORM EVENTS</p><h2>DDL 与活动</h2><p>DDL 按课程分类，活动记录不需要课程标签。</p></div><div className="button-row"><button className="secondary-button" onClick={() => downloadCalendar(filteredCalendar)}><FileDown size={16}/>导出 Excel</button><button className="primary-button" onClick={() => setShowForm(showForm === "ddl" ? "" : "ddl")}><Plus size={16}/>新增 DDL</button><button className="secondary-button" onClick={() => setShowForm(showForm === "activity" ? "" : "activity")}><Plus size={16}/>新增活动</button></div></div>
          {showForm === "ddl" && <CalendarForm kind="ddl" courses={data.courses} busy={busy} onClose={() => setShowForm("")} onSubmit={(form) => mutate({ action: "addCalendar", ...form }, "DDL 已加入宿舍日历")}/>}
          {showForm === "activity" && <CalendarForm kind="activity" courses={data.courses} busy={busy} onClose={() => setShowForm("")} onSubmit={(form) => mutate({ action: "addCalendar", ...form }, "活动已加入宿舍日历")}/>}
          <div className="calendar-layout"><section className="panel month-panel"><div className="month-toolbar"><div><p className="eyebrow">SHARED CALENDAR</p><h3>{month.replace("-", "年")}月</h3></div><div className="month-controls"><button aria-label="上个月" onClick={() => shiftMonth(month, -1, setMonth)}><ChevronLeft size={18}/></button><button className="month-today" onClick={() => { setMonth(shanghaiDate().slice(0, 7)); setSelectedDate(shanghaiDate()); }}>今天</button><button aria-label="下个月" onClick={() => shiftMonth(month, 1, setMonth)}><ChevronRight size={18}/></button></div></div>
            <div className="course-filters"><span>只看课程：</span>{data.courses.length ? data.courses.map((course) => <label key={course.id} className={`filter-pill ${courseFilter.includes(course.id) ? "selected" : ""}`}><input type="checkbox" checked={courseFilter.includes(course.id)} onChange={(e) => setCourseFilter((current) => e.target.checked ? [...current, course.id] : current.filter((id) => id !== course.id))}/>{course.name}</label>) : <small>暂无课程标签</small>}{courseFilter.length > 0 && <button className="clear-filter" onClick={() => setCourseFilter([])}>清除</button>}</div>
            <CalendarGrid month={month} selected={selectedDate} items={monthItems} onSelect={setSelectedDate}/>
            <div className="calendar-legend"><span><i className="legend-ddl"/>课程 DDL</span><span><i className="legend-event"/>宿舍活动</span><span className="legend-note">选择日期查看详情</span></div>
          </section><section className="panel agenda-panel"><div className="panel-head"><div><p className="eyebrow">AGENDA</p><h3>{humanDate(selectedDate)} <small>{selectedDate}</small></h3></div><span className="agenda-count">{selectedItems.length} 项</span></div>{selectedItems.length ? <div className="agenda-list">{selectedItems.map((item) => <CalendarRow key={item.id} item={item} memberId={data.member.id} admin={data.member.role === "admin"} onEdit={() => setEditing({ type: "calendar", item })} onDelete={() => remove("deleteCalendar", item.id, "删除这条日程？")} onCalendar={() => addToDeviceCalendar(item)}/>)}</div> : <div className="agenda-empty"><CalendarDays size={27}/><b>这一天还没有安排</b><span>可以新增一个 DDL 或宿舍活动。</span><button onClick={() => setShowForm("ddl")}><Plus size={15}/>添加事项</button></div>}</section></div>
          <section className="panel full-agenda"><div className="panel-head"><div><p className="eyebrow">VISIBLE TO YOU</p><h3>筛选后的日历记录</h3></div><span className="agenda-count">{filteredCalendar.length} 项</span></div>{filteredCalendar.length ? <div className="upcoming-list">{filteredCalendar.map((item) => <CalendarRow key={item.id} item={item} memberId={data.member.id} admin={data.member.role === "admin"} onEdit={() => setEditing({ type: "calendar", item })} onDelete={() => remove("deleteCalendar", item.id, "删除这条日程？")} onCalendar={() => addToDeviceCalendar(item)}/>)}</div> : <Empty title="还没有日历记录" text="新增 DDL 或宿舍活动，它们会出现在共享日历中。" action="新增 DDL" onClick={() => setShowForm("ddl")} />}</section>
        </section>}

        {tab === "duties" && <section className="page-content">
          <div className="section-intro"><div><p className="eyebrow">DAILY ROTATION</p><h2>线上值日表</h2><p>倒垃圾和扫地分别打卡，两项都完成后，当天值日才算完成。</p></div>{data.member.role === "admin" && <button className="primary-button" onClick={() => setShowForm(showForm === "duty" ? "" : "duty")}><Plus size={16}/>安排值日</button>}</div>
          {showForm === "duty" && <DutyForm members={data.members} busy={busy} onClose={() => setShowForm("")} onSubmit={(form) => mutate({ action: "addDuty", ...form }, "值日安排已添加")}/>}
          <div className="duty-summary"><div className="metric-icon green"><Check size={18}/></div><div><b>{doneCount} / {totalChecks}</b><span>已完成分项</span></div><div className="progress-track"><span style={{ width: `${totalChecks ? Math.round(doneCount / totalChecks * 100) : 0}%` }}/></div><small>{totalChecks ? Math.round(doneCount / totalChecks * 100) : 0}%</small></div>
          <div className="duty-list">{data.duties.map((d) => <article className={`duty-card ${d.garbageDone && d.sweepDone ? "duty-complete" : ""}`} key={d.id}><div className="duty-date"><span>{humanDate(d.date)}</span><small>{new Date(`${d.date}T00:00:00Z`).toLocaleDateString("zh-CN", { weekday: "short", timeZone: "UTC" })}</small></div><DutyTask label="倒垃圾" name={d.garbageName} done={!!d.garbageDone} me={d.garbageMemberId === data.member.id} admin={data.member.role === "admin"} onToggle={(done) => mutate({ action: "toggleDuty", id: d.id, field: "garbage", done }, done ? "倒垃圾已打卡" : "已撤销打卡")}/><DutyTask label="扫地" name={d.sweepName} done={!!d.sweepDone} me={d.sweepMemberId === data.member.id} admin={data.member.role === "admin"} onToggle={(done) => mutate({ action: "toggleDuty", id: d.id, field: "sweep", done }, done ? "扫地已打卡" : "已撤销打卡")}/>{data.member.role === "admin" && <IconButton label="编辑排班" onClick={() => setEditing({ type: "duty", item: d })}><Pencil size={15}/></IconButton>}{data.member.role === "admin" && <IconButton label="删除排班" danger onClick={() => remove("deleteDuty", d.id, "删除这天的值日安排？")}><Trash2 size={16}/></IconButton>}</article>)}{!data.duties.length && <Empty title="还没有值日安排" text={data.member.role === "admin" ? "添加日期并分别安排倒垃圾、扫地的成员。" : "等待管理员添加值日安排。"} action={data.member.role === "admin" ? "安排第一天值日" : undefined} onClick={() => setShowForm("duty")} />}</div>
        </section>}

        {tab === "sports" && <section className="page-content">
          <div className="section-intro"><div><p className="eyebrow">MOVE AT YOUR PACE</p><h2>运动打卡</h2><p>想记详细一些可以填运动类型和时长；只打卡也可以。没有排行榜。</p></div><button className="primary-button" onClick={() => setShowForm(showForm === "sport" ? "" : "sport")}><Plus size={16}/>运动打卡</button></div>
          {showForm === "sport" && <section className="form-panel"><div className="form-title"><div><p className="eyebrow">CHECK IN</p><h3>记录一次运动</h3></div><IconButton label="关闭" onClick={() => setShowForm("")}><X size={18}/></IconButton></div><form className="form-grid" onSubmit={(e) => { e.preventDefault(); const f = new FormData(e.currentTarget); void mutate({ action: "addSport", date: f.get("date"), activityType: f.get("activityType"), durationMinutes: f.get("durationMinutes"), isPublic: f.get("isPublic") === "on" }, "运动打卡已记录"); }}><label>日期<input name="date" type="date" defaultValue={shanghaiDate()} required/></label><label>运动类型 <span className="optional-label">选填</span><input name="activityType" placeholder="如：跑步、游泳、羽毛球" maxLength={80}/></label><label>时长（分钟） <span className="optional-label">选填</span><input name="durationMinutes" type="number" min="1" max="1440" placeholder="例如 30"/></label><label className="visibility-control"><input type="checkbox" name="isPublic" defaultChecked/>对全宿舍可见</label><div className="form-actions"><button className="primary-button" disabled={busy}><Check size={15}/>{busy ? "保存中…" : "完成打卡"}</button><span>取消勾选后，这条记录只有你能看到</span></div></form></section>}
          <div className="sports-list">{data.sports.map((sport) => <article key={sport.id} className="sport-card"><div className="sport-icon"><Dumbbell size={19}/></div><div className="sport-main"><h3>{sport.activityType || "完成一次运动"}</h3><p>{humanDate(sport.date)}{sport.durationMinutes ? ` · ${sport.durationMinutes} 分钟` : " · 只打卡"}</p></div><span className={`privacy-tag ${sport.isPublic ? "public" : "private"}`}>{sport.isPublic ? <Users size={13}/> : <LockKeyhole size={13}/>} {sport.isPublic ? "全宿舍可见" : "仅自己可见"}</span><span className="sport-owner">{sport.ownerId === data.member.id ? "我" : sport.owner}</span>{(sport.ownerId === data.member.id || data.member.role === "admin") && <IconButton label="编辑打卡" onClick={() => setEditing({ type: "sport", item: sport })}><Pencil size={15}/></IconButton>}{(sport.ownerId === data.member.id || data.member.role === "admin") && <IconButton label="删除打卡" danger onClick={() => remove("deleteSport", sport.id, "删除这条运动打卡？")}><Trash2 size={16}/></IconButton>}</article>)}{!data.sports.length && <Empty title="还没有运动打卡" text="不需要填很多内容，完成一次运动后点一下就能记下。" action="开始打卡" onClick={() => setShowForm("sport")} />}</div>
        </section>}

        {tab === "memories" && <section className="page-content">
          <div className="section-intro"><div><p className="eyebrow">LIFE TOGETHER</p><h2>宿舍风采</h2><p>把出游、生日、比赛和普通日子里的小事收进照片墙。</p></div><button className="primary-button" onClick={() => setShowForm(showForm === "memory" ? "" : "memory")}><Upload size={16}/>上传照片</button></div>
          {showForm === "memory" && <section className="form-panel"><div className="form-title"><div><p className="eyebrow">MEMORY WALL</p><h3>添加一张照片</h3></div><IconButton label="关闭" onClick={() => setShowForm("")}><X size={18}/></IconButton></div><form className="form-grid" onSubmit={(e) => { e.preventDefault(); const form = new FormData(e.currentTarget); form.set("type", "memory"); void upload(form, "照片已添加到宿舍风采"); }}><label>日期 <span className="optional-label">选填</span><input name="date" type="date"/></label><label>照片说明 <span className="optional-label">选填</span><input name="caption" placeholder="给这张照片留一句话" maxLength={500}/></label><label className="full-field file-drop"><Images size={18}/><span>选择照片</span><small>JPG、PNG、WebP 或 GIF · 单张不超过 25 MB</small><input name="file" type="file" accept="image/jpeg,image/png,image/webp,image/gif" required/></label><div className="form-actions"><button className="primary-button" disabled={busy}><Upload size={15}/>{busy ? "上传中…" : "添加到照片墙"}</button><span>你可以删除自己上传的照片</span></div></form></section>}
          <div className="memory-grid">{data.memories.map((memory) => <article className="memory-card" key={memory.id}><a className="memory-image" href={`/api/files?id=${encodeURIComponent(memory.id)}&inline=1`} target="_blank" rel="noreferrer"><img src={`/api/files?id=${encodeURIComponent(memory.id)}&inline=1`} alt={memory.caption || memory.fileName}/><span><Images size={16}/>查看照片</span></a><div className="memory-caption"><div><b>{memory.caption || "宿舍日常"}</b><small>{memory.date ? humanDate(memory.date) : humanDate(memory.createdAt.slice(0, 10))} · {memory.uploader}</small></div>{(memory.uploaderId === data.member.id || data.member.role === "admin") && <IconButton label="编辑照片说明" onClick={() => setEditing({ type: "memory", item: memory })}><Pencil size={15}/></IconButton>}{(memory.uploaderId === data.member.id || data.member.role === "admin") && <IconButton label="删除照片" danger onClick={() => remove("deleteMemory", memory.id, "删除这张照片？")}><Trash2 size={16}/></IconButton>}</div></article>)}{!data.memories.length && <div className="memory-empty"><div className="memory-empty-icon"><Images size={28}/></div><h3>共同的故事，从第一张照片开始</h3><p>上传一张照片，可选日期与说明。你能删除自己上传的照片，管理员可以协助管理照片墙。</p><button className="secondary-button" onClick={() => setShowForm("memory")}><Plus size={15}/>添加第一张照片</button></div>}</div>
        </section>}
      </>}
      {editing && <EditDialog edit={editing} courses={data?.courses ?? []} members={data?.members ?? []} busy={busy} onClose={() => setEditing(null)} onSubmit={saveEdit}/>}
      <footer className="site-footer"><span>511 宿舍共享空间</span><span>记录学习，也记录一起生活的日子。</span></footer>
      {notice && <div className="toast" role="status"><span><Check size={14}/></span>{notice}</div>}
    </main>
  </div>;
}

function CalendarForm({ kind, courses, busy, onClose, onSubmit }: { kind: "ddl" | "activity"; courses: Course[]; busy: boolean; onClose: () => void; onSubmit: (data: Record<string, unknown>) => void }) {
  return <section className="form-panel"><div className="form-title"><div><p className="eyebrow">{kind === "ddl" ? "COURSE DEADLINE" : "DORM EVENT"}</p><h3>{kind === "ddl" ? "新增课程 DDL" : "新增活动"}</h3></div><IconButton label="关闭" onClick={onClose}><X size={18}/></IconButton></div><form className="form-grid" onSubmit={(e) => { e.preventDefault(); const f = new FormData(e.currentTarget); onSubmit({ kind, title: f.get("title"), date: f.get("date"), courseId: f.get("courseId"), dueTime: f.get("dueTime"), startTime: f.get("startTime"), endTime: f.get("endTime"), location: f.get("location"), notes: f.get("notes"), isPublic: f.get("isPublic") === "on" }); }}>
    <label>事项名称<input name="title" maxLength={160} placeholder={kind === "ddl" ? "例如：文献综述初稿" : "例如：周末集体出游"} required/></label>
    {kind === "ddl" && <label>课程标签<select name="courseId" required defaultValue=""><option value="" disabled>选择课程</option>{courses.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></label>}
    <label>日期<input name="date" type="date" defaultValue={shanghaiDate()} required/></label>
    {kind === "ddl" ? <label>截止时间 <span className="optional-label">选填</span><input name="dueTime" type="time"/></label> : <><label>开始时间 <span className="optional-label">选填</span><input name="startTime" type="time"/></label><label>结束时间 <span className="optional-label">选填</span><input name="endTime" type="time"/></label><label>地点 <span className="optional-label">选填</span><input name="location" maxLength={160} placeholder="活动地点"/></label></>}
    <label className="full-field">备注 <span className="optional-label">选填</span><textarea name="notes" maxLength={1000} rows={2} placeholder="补充说明"/></label>
    {kind === "ddl" && <label className="visibility-control"><input type="checkbox" name="isPublic" defaultChecked/>全宿舍可见</label>}
    <div className="form-actions"><button className="primary-button" disabled={busy}><Plus size={15}/>{busy ? "保存中…" : "加入共享日历"}</button>{kind === "ddl" && <span>取消勾选后仅自己可见</span>}</div>
  </form></section>;
}

function DutyForm({ members, busy, onClose, onSubmit }: { members: Member[]; busy: boolean; onClose: () => void; onSubmit: (data: Record<string, unknown>) => void }) {
  return <section className="form-panel"><div className="form-title"><div><p className="eyebrow">DUTY ROTATION</p><h3>安排一天值日</h3></div><IconButton label="关闭" onClick={onClose}><X size={18}/></IconButton></div><form className="form-grid" onSubmit={(e) => { e.preventDefault(); const f = new FormData(e.currentTarget); onSubmit({ date: f.get("date"), garbageMemberId: f.get("garbage"), sweepMemberId: f.get("sweep") }); }}><label>值日日期<input name="date" type="date" defaultValue={shanghaiDate()} required/></label><label>倒垃圾<select name="garbage" required defaultValue=""><option value="" disabled>选择成员</option>{members.map((m) => <option key={m.id} value={m.id}>{m.displayName}</option>)}</select></label><label>扫地<select name="sweep" required defaultValue=""><option value="" disabled>选择成员</option>{members.map((m) => <option key={m.id} value={m.id}>{m.displayName}</option>)}</select></label><div className="form-actions"><button className="primary-button" disabled={busy}><Plus size={15}/>{busy ? "保存中…" : "保存排班"}</button><span>成员完成自己负责的项目后打卡</span></div></form></section>;
}

function EditDialog({ edit, courses, members, busy, onClose, onSubmit }: { edit: EditState; courses: Course[]; members: Member[]; busy: boolean; onClose: () => void; onSubmit: (data: Record<string, unknown>) => void }) {
  const title = { calendar: "编辑日历记录", resource: "编辑课程资料", duty: "修改值日安排", sport: "编辑运动打卡", memory: "编辑照片说明" }[edit.type];
  const send = (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault(); const form = new FormData(event.currentTarget);
    const value = (key: string) => form.get(key);
    if (edit.type === "calendar") onSubmit({ title: value("title"), date: value("date"), courseId: value("courseId"), dueTime: value("dueTime"), startTime: value("startTime"), endTime: value("endTime"), location: value("location"), notes: value("notes"), isPublic: value("isPublic") === "on" });
    if (edit.type === "resource") onSubmit({ courseId: value("courseId"), title: value("title") });
    if (edit.type === "duty") onSubmit({ date: value("date"), garbageMemberId: value("garbage"), sweepMemberId: value("sweep") });
    if (edit.type === "sport") onSubmit({ date: value("date"), activityType: value("activityType"), durationMinutes: value("durationMinutes"), isPublic: value("isPublic") === "on" });
    if (edit.type === "memory") onSubmit({ date: value("date"), caption: value("caption") });
  };
  return <div className="modal-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className="edit-modal" role="dialog" aria-modal="true" aria-labelledby="edit-heading"><div className="form-title"><div><p className="eyebrow">EDIT RECORD</p><h3 id="edit-heading">{title}</h3></div><IconButton label="关闭" onClick={onClose}><X size={18}/></IconButton></div><form className="form-grid" onSubmit={send}>
    {edit.type === "calendar" && <><label>事项名称<input name="title" defaultValue={edit.item.title} required maxLength={160}/></label>{edit.item.kind === "ddl" && <label>课程标签<select name="courseId" defaultValue={edit.item.courseId ?? ""} required>{courses.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></label>}<label>日期<input name="date" type="date" defaultValue={edit.item.date} required/></label>{edit.item.kind === "ddl" ? <label>截止时间 <span className="optional-label">选填</span><input name="dueTime" type="time" defaultValue={edit.item.dueTime ?? ""}/></label> : <><label>开始时间<input name="startTime" type="time" defaultValue={edit.item.startTime ?? ""}/></label><label>结束时间<input name="endTime" type="time" defaultValue={edit.item.endTime ?? ""}/></label><label>地点<input name="location" defaultValue={edit.item.location ?? ""} maxLength={160}/></label></>}<label className="full-field">备注<textarea name="notes" defaultValue={edit.item.notes ?? ""} rows={2} maxLength={1000}/></label>{edit.item.kind === "ddl" && <label className="visibility-control"><input type="checkbox" name="isPublic" defaultChecked={!!edit.item.isPublic}/>全宿舍可见</label>}</>}
    {edit.type === "resource" && <><label>课程标签<select name="courseId" defaultValue={edit.item.courseId} required>{courses.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></label><label>资料名称<input name="title" defaultValue={edit.item.title} required maxLength={160}/></label><p className="full-field edit-note">文件本身不变；这里只修改资料名称和课程归档。</p></>}
    {edit.type === "duty" && <><label>值日日期<input name="date" type="date" defaultValue={edit.item.date} required/></label><label>倒垃圾<select name="garbage" defaultValue={edit.item.garbageMemberId ?? ""} required>{members.map((m) => <option key={m.id} value={m.id}>{m.displayName}</option>)}</select></label><label>扫地<select name="sweep" defaultValue={edit.item.sweepMemberId ?? ""} required>{members.map((m) => <option key={m.id} value={m.id}>{m.displayName}</option>)}</select></label><p className="full-field edit-note">更换负责成员时，该项打卡会重置。</p></>}
    {edit.type === "sport" && <><label>日期<input name="date" type="date" defaultValue={edit.item.date} required/></label><label>运动类型 <span className="optional-label">选填</span><input name="activityType" defaultValue={edit.item.activityType ?? ""} maxLength={80}/></label><label>时长（分钟） <span className="optional-label">选填</span><input name="durationMinutes" type="number" min="1" max="1440" defaultValue={edit.item.durationMinutes ?? ""}/></label><label className="visibility-control"><input type="checkbox" name="isPublic" defaultChecked={!!edit.item.isPublic}/>对全宿舍可见</label></>}
    {edit.type === "memory" && <><label>日期 <span className="optional-label">选填</span><input name="date" type="date" defaultValue={edit.item.date ?? ""}/></label><label>照片说明 <span className="optional-label">选填</span><input name="caption" defaultValue={edit.item.caption ?? ""} maxLength={500}/></label></>}
    <div className="form-actions"><button className="primary-button" disabled={busy}><Check size={15}/>{busy ? "保存中…" : "保存修改"}</button><button className="secondary-button" type="button" onClick={onClose}>取消</button></div>
  </form></section></div>;
}

function DutyTask({ label, name, done, me, admin, onToggle }: { label: string; name: string | null; done: boolean; me: boolean; admin: boolean; onToggle: (done: boolean) => void }) {
  const canToggle = admin || me;
  return <div className={`duty-task ${done ? "task-done" : ""}`}><span className="task-check">{done ? <Check size={14}/> : <span/>}</span><span className="task-info"><b>{label}</b><small>{name || "待安排"}{done ? " · 已完成" : me ? " · 轮到你了" : ""}</small></span>{canToggle && <button className={`checkin-button ${done ? "checked" : ""}`} onClick={() => onToggle(!done)}>{done ? <><Check size={14}/>已打卡</> : "打卡"}</button>}</div>;
}

function CalendarRow({ item, memberId, admin, onEdit, onDelete, onCalendar }: { item: CalendarItem; memberId: string; admin: boolean; onEdit?: () => void; onDelete: () => void; onCalendar: () => void }) {
  return <article className={`calendar-row ${item.kind === "activity" ? "event-row" : ""}`}><div className="date-tile"><b>{item.date.slice(8)}</b><small>{item.date.slice(5, 7)}月</small></div><div className="calendar-row-main"><div className="calendar-row-top"><span className={`kind-tag ${item.kind}`}>{item.kind === "ddl" ? "DDL" : "活动"}</span>{item.course && <span className="course-pill">{item.course}</span>}{item.kind === "ddl" && !item.isPublic && <span className="privacy-tag private"><LockKeyhole size={12}/>仅自己</span>}</div><h4>{item.title}</h4><div className="row-details">{item.kind === "ddl" ? item.dueTime ? <span><Clock3 size={13}/>{item.dueTime} 截止</span> : <span><CalendarDays size={13}/>仅日期</span> : <span><Clock3 size={13}/>{item.startTime ?? "时间待定"}{item.endTime ? `–${item.endTime}` : ""}</span>}{item.location && <span><MapPin size={13}/>{item.location}</span>}{item.notes && <span className="row-note">{item.notes}</span>}</div><small className="row-owner">{item.ownerId === memberId ? "由我创建" : item.owner}</small></div><div className="calendar-row-actions"><button className="mini-action" title="添加到设备日历" onClick={onCalendar}><CalendarDays size={14}/><span>加到日历</span></button>{(item.ownerId === memberId || admin) && onEdit && <IconButton label="编辑日程" onClick={onEdit}><Pencil size={14}/></IconButton>}{(item.ownerId === memberId || admin) && <IconButton label="删除日程" danger onClick={onDelete}><Trash2 size={15}/></IconButton>}</div></article>;
}

function CalendarGrid({ month, selected, items, onSelect }: { month: string; selected: string; items: CalendarItem[]; onSelect: (date: string) => void }) {
  const [year, m] = month.split("-").map(Number); const first = new Date(Date.UTC(year, m - 1, 1)).getUTCDay(); const count = new Date(Date.UTC(year, m, 0)).getUTCDate();
  const cells: (number | null)[] = [...Array(first).fill(null), ...Array.from({ length: count }, (_, i) => i + 1)];
  while (cells.length % 7) cells.push(null);
  const days = ["日", "一", "二", "三", "四", "五", "六"];
  return <div className="calendar-grid"><div className="calendar-weekdays">{days.map((d) => <span key={d}>周{d}</span>)}</div><div className="calendar-cells">{cells.map((day, index) => { const date = day ? `${year}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}` : ""; const dayItems = items.filter((item) => item.date === date); return <button key={index} disabled={!day} onClick={() => day && onSelect(date)} className={`calendar-cell ${date === selected ? "cell-selected" : ""} ${date === shanghaiDate() ? "cell-today" : ""}`}><span>{day ?? ""}</span>{dayItems.slice(0, 3).map((item) => <i key={item.id} className={`cell-event ${item.kind}`} title={item.title}>{item.title}</i>)}{dayItems.length > 3 && <small>+{dayItems.length - 3} 项</small>}<b className="cell-mobile-dots">{dayItems.length ? "•" : ""}</b></button>; })}</div></div>;
}

function Empty({ title, text, action, onClick }: { title: string; text: string; action?: string; onClick?: () => void }) {
  return <div className="empty-state"><span className="empty-mark">＋</span><b>{title}</b><p>{text}</p>{action && onClick && <button className="text-button" onClick={onClick}>{action} <ChevronRight size={14}/></button>}</div>;
}

function shiftMonth(value: string, amount: number, set: (value: string) => void) { const [y, m] = value.split("-").map(Number); const next = new Date(Date.UTC(y, m - 1 + amount, 1)); set(`${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, "0")}`); }
