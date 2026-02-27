import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "md2doc",
  description: "Markdown to DOCX with fixed preset"
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="zh-CN">
      <body style={{ margin: 0, fontFamily: "system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
        {children}
      </body>
    </html>
  );
}