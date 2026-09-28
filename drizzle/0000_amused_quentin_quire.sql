CREATE TABLE `calendar_items` (
	`id` text PRIMARY KEY NOT NULL,
	`kind` text NOT NULL,
	`title` text NOT NULL,
	`course_id` text,
	`item_date` text NOT NULL,
	`due_time` text,
	`start_time` text,
	`end_time` text,
	`location` text,
	`notes` text,
	`is_public` integer DEFAULT 1 NOT NULL,
	`owner_id` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`course_id`) REFERENCES `courses`(`id`) ON UPDATE no action ON DELETE no action,
	FOREIGN KEY (`owner_id`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action
);
--> statement-breakpoint
CREATE INDEX `calendar_date_kind` ON `calendar_items` (`item_date`,`kind`);--> statement-breakpoint
CREATE TABLE `courses` (
	`id` text PRIMARY KEY NOT NULL,
	`name` text NOT NULL,
	`created_by` text NOT NULL,
	`created_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `courses_name_unique` ON `courses` (`name`);--> statement-breakpoint
CREATE TABLE `duty_rota` (
	`id` text PRIMARY KEY NOT NULL,
	`duty_date` text NOT NULL,
	`garbage_member_id` text,
	`sweep_member_id` text,
	`garbage_done` integer DEFAULT 0 NOT NULL,
	`sweep_done` integer DEFAULT 0 NOT NULL,
	`garbage_done_by` text,
	`sweep_done_by` text,
	`updated_at` text NOT NULL,
	FOREIGN KEY (`garbage_member_id`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action,
	FOREIGN KEY (`sweep_member_id`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action,
	FOREIGN KEY (`garbage_done_by`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action,
	FOREIGN KEY (`sweep_done_by`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action
);
--> statement-breakpoint
CREATE UNIQUE INDEX `duty_rota_duty_date_unique` ON `duty_rota` (`duty_date`);--> statement-breakpoint
CREATE INDEX `duty_rota_date` ON `duty_rota` (`duty_date`);--> statement-breakpoint
CREATE TABLE `members` (
	`id` text PRIMARY KEY NOT NULL,
	`email` text NOT NULL,
	`display_name` text NOT NULL,
	`role` text DEFAULT 'member' NOT NULL,
	`joined_at` text NOT NULL
);
--> statement-breakpoint
CREATE UNIQUE INDEX `members_one_admin` ON `members` (`role`) WHERE "members"."role" = 'admin';--> statement-breakpoint
CREATE TABLE `memories` (
	`id` text PRIMARY KEY NOT NULL,
	`caption` text,
	`memory_date` text,
	`file_name` text NOT NULL,
	`object_key` text NOT NULL,
	`content_type` text NOT NULL,
	`size` integer NOT NULL,
	`uploader_id` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`uploader_id`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action
);
--> statement-breakpoint
CREATE UNIQUE INDEX `memories_object_key_unique` ON `memories` (`object_key`);--> statement-breakpoint
CREATE INDEX `memories_date` ON `memories` (`memory_date`);--> statement-breakpoint
CREATE TABLE `resources` (
	`id` text PRIMARY KEY NOT NULL,
	`course_id` text NOT NULL,
	`title` text NOT NULL,
	`file_name` text NOT NULL,
	`object_key` text NOT NULL,
	`content_type` text NOT NULL,
	`size` integer NOT NULL,
	`uploader_id` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`course_id`) REFERENCES `courses`(`id`) ON UPDATE no action ON DELETE no action,
	FOREIGN KEY (`uploader_id`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action
);
--> statement-breakpoint
CREATE UNIQUE INDEX `resources_object_key_unique` ON `resources` (`object_key`);--> statement-breakpoint
CREATE INDEX `resources_course_created` ON `resources` (`course_id`,`created_at`);--> statement-breakpoint
CREATE TABLE `sport_logs` (
	`id` text PRIMARY KEY NOT NULL,
	`activity_type` text,
	`duration_minutes` integer,
	`activity_date` text NOT NULL,
	`is_public` integer DEFAULT 1 NOT NULL,
	`owner_id` text NOT NULL,
	`created_at` text NOT NULL,
	FOREIGN KEY (`owner_id`) REFERENCES `members`(`id`) ON UPDATE no action ON DELETE no action
);
--> statement-breakpoint
CREATE INDEX `sport_logs_date` ON `sport_logs` (`activity_date`);