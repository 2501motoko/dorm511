import { env } from "cloudflare:workers";
import { getChatGPTUser } from "../../chatgpt-auth";

type Member = { id: string; email: string; displayName: string; role: "admin" | "member" };
const now = () => new Date().toISOString();
const json = (data: unknown, status = 200) => Response.json(data, { status });
const bad = (message: string, status = 400) => json({ error: message }, status);

async function currentMember(): Promise<Member | Response> {
  const identity = await getChatGPTUser();
  if (!identity) return bad("请先登录后再使用宿舍网站。", 401);
  const db = env.DB;
  if (!db) return bad("共享数据暂时不可用，请稍后重试。", 503);
  try {
    await db.prepare(`INSERT OR IGNORE INTO members (id,email,display_name,role,joined_at)
      SELECT ?,?,?,CASE WHEN NOT EXISTS (SELECT 1 FROM members WHERE role='admin') THEN 'admin' ELSE 'member' END,?`)
      .bind(identity.userId, identity.email, identity.displayName, now()).run();
  } catch {
    await db.prepare(`INSERT OR IGNORE INTO members (id,email,display_name,role,joined_at) VALUES (?,?,?,'member',?)`)
      .bind(identity.userId, identity.email, identity.displayName, now()).run();
  }
  await db.prepare(`INSERT OR IGNORE INTO members (id,email,display_name,role,joined_at) VALUES (?,?,?,'member',?)`)
    .bind(identity.userId, identity.email, identity.displayName, now()).run();
  const row = await db.prepare("SELECT id,email,display_name,role FROM members WHERE id=?")
    .bind(identity.userId).first<{ id: string; email: string; display_name: string; role: "admin" | "member" }>();
  if (!row) return bad("无法读取成员信息。", 503);
  return { id: row.id, email: row.email, displayName: row.display_name, role: row.role };
}

function isMember(value: Member | Response): value is Member { return !(value instanceof Response); }

export async function GET() {
  const member = await currentMember();
  if (!isMember(member)) return member;
  try {
    const db = env.DB;
    const [courses, resources, calendar, duties, sports, memories, members] = await Promise.all([
      db.prepare("SELECT id,name FROM courses ORDER BY name COLLATE NOCASE").all(),
      db.prepare(`SELECT r.id,r.course_id AS courseId,c.name AS course,r.title,r.file_name AS fileName,r.content_type AS contentType,r.size,r.uploader_id AS uploaderId,m.display_name AS uploader,r.created_at AS createdAt
        FROM resources r JOIN courses c ON c.id=r.course_id JOIN members m ON m.id=r.uploader_id ORDER BY r.created_at DESC`).all(),
      db.prepare(`SELECT e.id,e.kind,e.title,e.course_id AS courseId,c.name AS course,e.item_date AS date,e.due_time AS dueTime,e.start_time AS startTime,e.end_time AS endTime,e.location,e.notes,e.is_public AS isPublic,e.owner_id AS ownerId,m.display_name AS owner
        FROM calendar_items e LEFT JOIN courses c ON c.id=e.course_id JOIN members m ON m.id=e.owner_id
        WHERE e.is_public=1 OR e.owner_id=? ORDER BY e.item_date,e.due_time,e.start_time`).bind(member.id).all(),
      db.prepare(`SELECT d.id,d.duty_date AS date,d.garbage_member_id AS garbageMemberId,d.sweep_member_id AS sweepMemberId,
        d.garbage_done AS garbageDone,d.sweep_done AS sweepDone,d.garbage_done_by AS garbageDoneBy,d.sweep_done_by AS sweepDoneBy,
        g.display_name AS garbageName,s.display_name AS sweepName FROM duty_rota d
        LEFT JOIN members g ON g.id=d.garbage_member_id LEFT JOIN members s ON s.id=d.sweep_member_id ORDER BY d.duty_date`).all(),
      db.prepare(`SELECT x.id,x.activity_type AS activityType,x.duration_minutes AS durationMinutes,x.activity_date AS date,x.is_public AS isPublic,x.owner_id AS ownerId,m.display_name AS owner
        FROM sport_logs x JOIN members m ON m.id=x.owner_id WHERE x.is_public=1 OR x.owner_id=? ORDER BY x.activity_date DESC`).bind(member.id).all(),
      db.prepare(`SELECT p.id,p.caption,p.memory_date AS date,p.file_name AS fileName,p.content_type AS contentType,p.size,p.uploader_id AS uploaderId,m.display_name AS uploader,p.created_at AS createdAt
        FROM memories p JOIN members m ON m.id=p.uploader_id ORDER BY COALESCE(p.memory_date,p.created_at) DESC`).all(),
      db.prepare("SELECT id,display_name AS displayName,role FROM members ORDER BY display_name COLLATE NOCASE").all(),
    ]);
    return json({ member, courses: courses.results, resources: resources.results, calendar: calendar.results, duties: duties.results,
      sports: sports.results, memories: memories.results, members: members.results });
  } catch (error) {
    console.error("Data load failed", error);
    return bad("共享数据暂时无法读取，请稍后重试。", 503);
  }
}

export async function POST(request: Request) {
  const member = await currentMember();
  if (!isMember(member)) return member;
  const db = env.DB;
  let body: Record<string, unknown>;
  try { body = await request.json() as Record<string, unknown>; }
  catch { return bad("请求格式不正确。"); }
  const action = String(body.action ?? "");
  const str = (key: string, max = 500) => typeof body[key] === "string" ? String(body[key]).trim().slice(0, max) : "";
  const needAdmin = () => member.role === "admin" ? null : bad("这项操作仅管理员可用。", 403);
  try {
    if (action === "addCourse") {
      const denied = needAdmin(); if (denied) return denied;
      const name = str("name", 80); if (!name) return bad("请填写课程名称。");
      const id = crypto.randomUUID();
      await db.prepare("INSERT INTO courses(id,name,created_by,created_at) VALUES(?,?,?,?)")
        .bind(id, name, member.id, now()).run();
      return json({ id, name }, 201);
    }
    if (action === "deleteCourse") {
      const denied = needAdmin(); if (denied) return denied;
      const id = str("id", 80);
      const used = await db.prepare("SELECT 1 AS found FROM resources WHERE course_id=? UNION SELECT 1 FROM calendar_items WHERE course_id=? LIMIT 1")
        .bind(id, id).first();
      if (used) return bad("这门课已有资料或 DDL，不能删除。", 409);
      await db.prepare("DELETE FROM courses WHERE id=?").bind(id).run();
      return json({ ok: true });
    }
    if (action === "addCalendar") {
      const kind = str("kind", 16); const title = str("title", 160); const date = str("date", 10);
      if (!title || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad("请填写事项名称和日期。");
      const courseId = str("courseId", 80) || null;
      const dueTime = str("dueTime", 5) || null;
      const startTime = str("startTime", 5) || null; const endTime = str("endTime", 5) || null;
      if (kind === "ddl" && (!courseId || !await db.prepare("SELECT 1 FROM courses WHERE id=?").bind(courseId).first())) return bad("请选择课程标签。");
      if (kind !== "ddl" && kind !== "activity") return bad("事项类型不正确。");
      if (kind === "activity" && startTime && endTime && endTime < startTime) return bad("结束时间不能早于开始时间。");
      const id = crypto.randomUUID(); const isPublic = kind === "activity" ? 1 : (body.isPublic ? 1 : 0);
      await db.prepare(`INSERT INTO calendar_items(id,kind,title,course_id,item_date,due_time,start_time,end_time,location,notes,is_public,owner_id,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)`).bind(id, kind, title, kind === "ddl" ? courseId : null, date,
          kind === "ddl" ? dueTime : null, kind === "activity" ? startTime : null, kind === "activity" ? endTime : null,
          kind === "activity" ? str("location", 160) || null : null, str("notes", 1000) || null, isPublic, member.id, now()).run();
      return json({ id }, 201);
    }
    if (action === "deleteCalendar") {
      const id = str("id", 80);
      const row = await db.prepare("SELECT owner_id FROM calendar_items WHERE id=?").bind(id).first<{ owner_id: string }>();
      if (!row) return bad("这条日历记录不存在。", 404);
      if (row.owner_id !== member.id && member.role !== "admin") return bad("只能删除自己创建的记录。", 403);
      await db.prepare("DELETE FROM calendar_items WHERE id=?").bind(id).run();
      return json({ ok: true });
    }
    if (action === "editCalendar") {
      const id = str("id", 80); const row = await db.prepare("SELECT owner_id,kind FROM calendar_items WHERE id=?").bind(id).first<{ owner_id: string; kind: "ddl" | "activity" }>();
      if (!row) return bad("日历记录不存在。", 404);
      if (row.owner_id !== member.id && member.role !== "admin") return bad("只能编辑自己创建的记录。", 403);
      const title = str("title", 160), date = str("date", 10), kind = row.kind;
      if (!title || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad("请填写事项名称和日期。");
      const courseId = str("courseId", 80) || null, dueTime = str("dueTime", 5) || null;
      const startTime = str("startTime", 5) || null, endTime = str("endTime", 5) || null;
      if (kind === "ddl" && (!courseId || !await db.prepare("SELECT 1 FROM courses WHERE id=?").bind(courseId).first())) return bad("请选择课程标签。");
      if (kind === "activity" && startTime && endTime && endTime < startTime) return bad("结束时间不能早于开始时间。");
      await db.prepare(`UPDATE calendar_items SET title=?,course_id=?,item_date=?,due_time=?,start_time=?,end_time=?,location=?,notes=?,is_public=? WHERE id=?`)
        .bind(title, kind === "ddl" ? courseId : null, date, kind === "ddl" ? dueTime : null,
          kind === "activity" ? startTime : null, kind === "activity" ? endTime : null,
          kind === "activity" ? str("location", 160) || null : null, str("notes", 1000) || null,
          kind === "activity" || body.isPublic ? 1 : 0, id).run();
      return json({ ok: true });
    }
    if (action === "addDuty") {
      const denied = needAdmin(); if (denied) return denied;
      const date = str("date", 10), garbage = str("garbageMemberId", 100), sweep = str("sweepMemberId", 100);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !garbage || !sweep) return bad("请选择日期以及倒垃圾、扫地的值日成员。");
      const id = crypto.randomUUID();
      await db.prepare("INSERT INTO duty_rota(id,duty_date,garbage_member_id,sweep_member_id,updated_at) VALUES(?,?,?,?,?)")
        .bind(id, date, garbage, sweep, now()).run();
      return json({ id }, 201);
    }
    if (action === "toggleDuty") {
      const field = str("field", 20); const done = body.done ? 1 : 0; const id = str("id", 80);
      if (field !== "garbage" && field !== "sweep") return bad("打卡项目不正确。");
      const col = field === "garbage" ? "garbage" : "sweep";
      const row = await db.prepare(`SELECT ${col}_member_id AS assigned FROM duty_rota WHERE id=?`).bind(id).first<{ assigned: string | null }>();
      if (!row) return bad("没有找到这条值日安排。", 404);
      if (member.role !== "admin" && row.assigned !== member.id) return bad("只有对应值日成员可以打卡。", 403);
      await db.prepare(`UPDATE duty_rota SET ${col}_done=?,${col}_done_by=?,updated_at=? WHERE id=?`)
        .bind(done, done ? member.id : null, now(), id).run();
      return json({ ok: true });
    }
    if (action === "deleteDuty") {
      const denied = needAdmin(); if (denied) return denied;
      await db.prepare("DELETE FROM duty_rota WHERE id=?").bind(str("id", 80)).run();
      return json({ ok: true });
    }
    if (action === "editDuty") {
      const denied = needAdmin(); if (denied) return denied;
      const id = str("id", 80), date = str("date", 10), garbage = str("garbageMemberId", 100), sweep = str("sweepMemberId", 100);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !garbage || !sweep) return bad("请选择日期以及两项值日成员。");
      const current = await db.prepare("SELECT garbage_member_id,sweep_member_id FROM duty_rota WHERE id=?").bind(id).first<{ garbage_member_id: string | null; sweep_member_id: string | null }>();
      if (!current) return bad("没有找到这条值日安排。", 404);
      const garbageChanged = current.garbage_member_id !== garbage, sweepChanged = current.sweep_member_id !== sweep;
      await db.prepare(`UPDATE duty_rota SET duty_date=?,garbage_member_id=?,sweep_member_id=?,
        garbage_done=CASE WHEN ? THEN 0 ELSE garbage_done END,garbage_done_by=CASE WHEN ? THEN NULL ELSE garbage_done_by END,
        sweep_done=CASE WHEN ? THEN 0 ELSE sweep_done END,sweep_done_by=CASE WHEN ? THEN NULL ELSE sweep_done_by END,updated_at=? WHERE id=?`)
        .bind(date, garbage, sweep, garbageChanged ? 1 : 0, garbageChanged ? 1 : 0, sweepChanged ? 1 : 0, sweepChanged ? 1 : 0, now(), id).run();
      return json({ ok: true });
    }
    if (action === "addSport") {
      const date = str("date", 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad("请填写运动日期。");
      const type = str("activityType", 80) || null;
      const rawDuration = body.durationMinutes;
      const duration = rawDuration === "" || rawDuration == null ? null : Number(rawDuration);
      if (duration !== null && (!Number.isInteger(duration) || duration < 1 || duration > 1440)) return bad("运动时长需为 1 到 1440 分钟。");
      const id = crypto.randomUUID();
      await db.prepare("INSERT INTO sport_logs(id,activity_type,duration_minutes,activity_date,is_public,owner_id,created_at) VALUES(?,?,?,?,?,?,?)")
        .bind(id, type, duration, date, body.isPublic ? 1 : 0, member.id, now()).run();
      return json({ id }, 201);
    }
    if (action === "deleteSport") {
      const id = str("id", 80); const row = await db.prepare("SELECT owner_id FROM sport_logs WHERE id=?").bind(id).first<{ owner_id: string }>();
      if (!row) return bad("没有找到这条运动打卡。", 404);
      if (row.owner_id !== member.id && member.role !== "admin") return bad("只能删除自己的打卡。", 403);
      await db.prepare("DELETE FROM sport_logs WHERE id=?").bind(id).run(); return json({ ok: true });
    }
    if (action === "editSport") {
      const id = str("id", 80); const row = await db.prepare("SELECT owner_id FROM sport_logs WHERE id=?").bind(id).first<{ owner_id: string }>();
      if (!row) return bad("没有找到这条运动打卡。", 404);
      if (row.owner_id !== member.id && member.role !== "admin") return bad("只能编辑自己的打卡。", 403);
      const date = str("date", 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad("请填写运动日期。");
      const type = str("activityType", 80) || null;
      const rawDuration = body.durationMinutes, duration = rawDuration === "" || rawDuration == null ? null : Number(rawDuration);
      if (duration !== null && (!Number.isInteger(duration) || duration < 1 || duration > 1440)) return bad("运动时长需为 1 到 1440 分钟。");
      await db.prepare("UPDATE sport_logs SET activity_type=?,duration_minutes=?,activity_date=?,is_public=? WHERE id=?")
        .bind(type, duration, date, body.isPublic ? 1 : 0, id).run();
      return json({ ok: true });
    }
    if (action === "editResource") {
      const id = str("id", 80); const row = await db.prepare("SELECT uploader_id FROM resources WHERE id=?").bind(id).first<{ uploader_id: string }>();
      if (!row) return bad("资料不存在。", 404);
      if (row.uploader_id !== member.id && member.role !== "admin") return bad("只能编辑自己上传的资料。", 403);
      const courseId = str("courseId", 80), title = str("title", 160);
      if (!title || !await db.prepare("SELECT 1 FROM courses WHERE id=?").bind(courseId).first()) return bad("请填写资料名称并选择课程。");
      await db.prepare("UPDATE resources SET course_id=?,title=? WHERE id=?").bind(courseId, title, id).run();
      return json({ ok: true });
    }
    if (action === "editMemory") {
      const id = str("id", 80); const row = await db.prepare("SELECT uploader_id FROM memories WHERE id=?").bind(id).first<{ uploader_id: string }>();
      if (!row) return bad("照片不存在。", 404);
      if (row.uploader_id !== member.id && member.role !== "admin") return bad("只能编辑自己上传的照片。", 403);
      const date = str("date", 10) || null;
      if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) return bad("日期格式不正确。");
      await db.prepare("UPDATE memories SET caption=?,memory_date=? WHERE id=?").bind(str("caption", 500) || null, date, id).run();
      return json({ ok: true });
    }
    if (action === "deleteResource") {
      const id = str("id", 80); const row = await db.prepare("SELECT object_key,uploader_id FROM resources WHERE id=?").bind(id).first<{ object_key: string; uploader_id: string }>();
      if (!row) return bad("资料不存在。", 404);
      if (row.uploader_id !== member.id && member.role !== "admin") return bad("只能删除自己上传的资料。", 403);
      await env.BUCKET.delete(row.object_key); await db.prepare("DELETE FROM resources WHERE id=?").bind(id).run(); return json({ ok: true });
    }
    if (action === "deleteMemory") {
      const id = str("id", 80); const row = await db.prepare("SELECT object_key,uploader_id FROM memories WHERE id=?").bind(id).first<{ object_key: string; uploader_id: string }>();
      if (!row) return bad("照片不存在。", 404);
      if (row.uploader_id !== member.id && member.role !== "admin") return bad("只能删除自己上传的照片。", 403);
      await env.BUCKET.delete(row.object_key); await db.prepare("DELETE FROM memories WHERE id=?").bind(id).run(); return json({ ok: true });
    }
    return bad("不支持的操作。");
  } catch (error) {
    console.error("Data write failed", action, error);
    if (String(error).includes("UNIQUE constraint failed")) return bad("已有相同名称或同一天的记录。", 409);
    return bad("保存失败，请检查填写内容后重试。", 503);
  }
}
