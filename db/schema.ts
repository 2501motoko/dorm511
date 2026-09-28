import { sql } from "drizzle-orm";
import { index, integer, sqliteTable, text, uniqueIndex } from "drizzle-orm/sqlite-core";

export const members = sqliteTable(
  "members",
  {
    id: text("id").primaryKey(),
    email: text("email").notNull(),
    displayName: text("display_name").notNull(),
    role: text("role", { enum: ["admin", "member"] }).notNull().default("member"),
    joinedAt: text("joined_at").notNull(),
  },
  (table) => [uniqueIndex("members_one_admin").on(table.role).where(sql`${table.role} = 'admin'`)],
);

export const courses = sqliteTable("courses", {
  id: text("id").primaryKey(),
  name: text("name").notNull().unique(),
  createdBy: text("created_by").notNull(),
  createdAt: text("created_at").notNull(),
});

export const resources = sqliteTable(
  "resources",
  {
    id: text("id").primaryKey(),
    courseId: text("course_id").notNull().references(() => courses.id),
    title: text("title").notNull(),
    fileName: text("file_name").notNull(),
    objectKey: text("object_key").notNull().unique(),
    contentType: text("content_type").notNull(),
    size: integer("size").notNull(),
    uploaderId: text("uploader_id").notNull().references(() => members.id),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("resources_course_created").on(table.courseId, table.createdAt)],
);

export const calendarItems = sqliteTable(
  "calendar_items",
  {
    id: text("id").primaryKey(),
    kind: text("kind", { enum: ["ddl", "activity"] }).notNull(),
    title: text("title").notNull(),
    courseId: text("course_id").references(() => courses.id),
    itemDate: text("item_date").notNull(),
    dueTime: text("due_time"),
    startTime: text("start_time"),
    endTime: text("end_time"),
    location: text("location"),
    notes: text("notes"),
    isPublic: integer("is_public").notNull().default(1),
    ownerId: text("owner_id").notNull().references(() => members.id),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("calendar_date_kind").on(table.itemDate, table.kind)],
);

export const dutyRota = sqliteTable(
  "duty_rota",
  {
    id: text("id").primaryKey(),
    dutyDate: text("duty_date").notNull().unique(),
    garbageMemberId: text("garbage_member_id").references(() => members.id),
    sweepMemberId: text("sweep_member_id").references(() => members.id),
    garbageDone: integer("garbage_done").notNull().default(0),
    sweepDone: integer("sweep_done").notNull().default(0),
    garbageDoneBy: text("garbage_done_by").references(() => members.id),
    sweepDoneBy: text("sweep_done_by").references(() => members.id),
    updatedAt: text("updated_at").notNull(),
  },
  (table) => [index("duty_rota_date").on(table.dutyDate)],
);

export const sportLogs = sqliteTable(
  "sport_logs",
  {
    id: text("id").primaryKey(),
    activityType: text("activity_type"),
    durationMinutes: integer("duration_minutes"),
    activityDate: text("activity_date").notNull(),
    isPublic: integer("is_public").notNull().default(1),
    ownerId: text("owner_id").notNull().references(() => members.id),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("sport_logs_date").on(table.activityDate)],
);

export const memories = sqliteTable(
  "memories",
  {
    id: text("id").primaryKey(),
    caption: text("caption"),
    memoryDate: text("memory_date"),
    fileName: text("file_name").notNull(),
    objectKey: text("object_key").notNull().unique(),
    contentType: text("content_type").notNull(),
    size: integer("size").notNull(),
    uploaderId: text("uploader_id").notNull().references(() => members.id),
    createdAt: text("created_at").notNull(),
  },
  (table) => [index("memories_date").on(table.memoryDate)],
);
