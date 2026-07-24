import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "과학전람회 엑셀 도구",
  description:
    "엑셀 행 분리와 과학전람회 상장용 명단 생성을 브라우저에서 처리하는 도구",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
