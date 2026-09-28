import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "511 社会学宿舍共享空间",
  description: "511 宿舍的课程资料、DDL 与活动日历、值日打卡、运动记录和宿舍照片墙。",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body className="antialiased">{children}</body>
    </html>
  );
}
