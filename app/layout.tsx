import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "md2doc",
  description: "Markdown to DOCX with fixed preset"
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
