import { env } from "cloudflare:workers";
import { getChatGPTUser } from "../../chatgpt-auth";

const MAX_BYTES = 25 * 1024 * 1024;
const supportedDocs = new Set(["pdf", "doc", "docx", "ppt", "pptx", "xls", "xlsx", "csv", "txt", "md", "zip"]);

async function signedMember() {
  const user = await getChatGPTUser();
  if (!user) return null;
  await env.DB.prepare(`INSERT OR IGNORE INTO members (id,email,display_name,role,joined_at)
    SELECT ?,?,?,CASE WHEN NOT EXISTS (SELECT 1 FROM members WHERE role='admin') THEN 'admin' ELSE 'member' END,?`)
    .bind(user.userId, user.email, user.displayName, new Date().toISOString()).run();
  await env.DB.prepare(`INSERT OR IGNORE INTO members (id,email,display_name,role,joined_at) VALUES (?,?,?,'member',?)`)
    .bind(user.userId, user.email, user.displayName, new Date().toISOString()).run();
  return user;
}

export async function POST(request: Request) {
  const user = await signedMember();
  if (!user) return Response.json({ error: "请先登录。" }, { status: 401 });
  if (!env.BUCKET) return Response.json({ error: "文件存储暂不可用。" }, { status: 503 });
  const form = await request.formData();
  const type = String(form.get("type") ?? "");
  const file = form.get("file");
  if (!(file instanceof File) || file.size === 0) return Response.json({ error: "请选择一个文件。" }, { status: 400 });
  if (file.size > MAX_BYTES) return Response.json({ error: "单个文件不能超过 25 MB。" }, { status: 413 });
  const original = file.name.replace(/[\\/\r\n\u0000]/g, "_").slice(0, 180);
  const ext = original.split(".").pop()?.toLowerCase() ?? "";
  if (type === "resource") {
    if (!supportedDocs.has(ext)) return Response.json({ error: "资料支持 PDF、Office 文档、CSV、TXT、Markdown 和 ZIP。" }, { status: 415 });
    const courseId = String(form.get("courseId") ?? "");
    const title = String(form.get("title") ?? original).trim().slice(0, 160);
    if (!courseId || !title) return Response.json({ error: "请填写资料名称并选择课程。" }, { status: 400 });
    const course = await env.DB.prepare("SELECT id FROM courses WHERE id=?").bind(courseId).first();
    if (!course) return Response.json({ error: "请选择有效课程标签。" }, { status: 400 });
    const id = crypto.randomUUID(), key = `resources/${id}`;
    await env.BUCKET.put(key, file.stream(), { httpMetadata: { contentType: file.type || "application/octet-stream" }, customMetadata: { fileName: original } });
    try {
      await env.DB.prepare("INSERT INTO resources(id,course_id,title,file_name,object_key,content_type,size,uploader_id,created_at) VALUES(?,?,?,?,?,?,?,?,?)")
        .bind(id, courseId, title, original, key, file.type || "application/octet-stream", file.size, user.userId, new Date().toISOString()).run();
    } catch (error) { await env.BUCKET.delete(key); throw error; }
    return Response.json({ id }, { status: 201 });
  }
  if (type === "memory") {
    if (!new Set(["jpg", "jpeg", "png", "webp", "gif"]).has(ext)) return Response.json({ error: "照片支持 JPG、PNG、WebP 或 GIF。" }, { status: 415 });
    const caption = String(form.get("caption") ?? "").trim().slice(0, 500) || null;
    const date = String(form.get("date") ?? "").trim() || null;
    if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) return Response.json({ error: "日期格式不正确。" }, { status: 400 });
    const id = crypto.randomUUID(), key = `memories/${id}`;
    await env.BUCKET.put(key, file.stream(), { httpMetadata: { contentType: file.type || "image/jpeg" }, customMetadata: { fileName: original } });
    try {
      await env.DB.prepare("INSERT INTO memories(id,caption,memory_date,file_name,object_key,content_type,size,uploader_id,created_at) VALUES(?,?,?,?,?,?,?,?,?)")
        .bind(id, caption, date, original, key, file.type || "image/jpeg", file.size, user.userId, new Date().toISOString()).run();
    } catch (error) { await env.BUCKET.delete(key); throw error; }
    return Response.json({ id }, { status: 201 });
  }
  return Response.json({ error: "上传类型不正确。" }, { status: 400 });
}

export async function GET(request: Request) {
  const user = await signedMember();
  if (!user) return Response.json({ error: "请先登录。" }, { status: 401 });
  const id = new URL(request.url).searchParams.get("id");
  if (!id || !env.BUCKET) return Response.json({ error: "文件不存在。" }, { status: 404 });
  const resource = await env.DB.prepare("SELECT file_name AS name,object_key AS key,content_type AS type FROM resources WHERE id=?").bind(id).first<{ name: string; key: string; type: string }>();
  const memory = resource ? null : await env.DB.prepare("SELECT file_name AS name,object_key AS key,content_type AS type FROM memories WHERE id=?").bind(id).first<{ name: string; key: string; type: string }>();
  const item = resource ?? memory;
  if (!item) return Response.json({ error: "文件不存在。" }, { status: 404 });
  const object = await env.BUCKET.get(item.key);
  if (!object) return Response.json({ error: "文件已不可用。" }, { status: 404 });
  const inline = new URL(request.url).searchParams.get("inline") === "1" && item.type.startsWith("image/");
  const safeName = encodeURIComponent(item.name).replace(/'/g, "%27");
  return new Response(object.body, { headers: {
    "Content-Type": item.type,
    "Content-Length": String(object.size),
    "Content-Disposition": `${inline ? "inline" : "attachment"}; filename*=UTF-8''${safeName}`,
    "Cache-Control": "private, max-age=300",
    "X-Content-Type-Options": "nosniff",
  } });
}
