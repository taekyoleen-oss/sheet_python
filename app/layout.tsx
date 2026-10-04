import type { Metadata } from "next";
import localFont from "next/font/local";
import { Fraunces, JetBrains_Mono } from "next/font/google";
import { Toaster } from "@/components/ui/sonner";
import "./globals.css";

const pretendard = localFont({
  src: "./fonts/PretendardVariable.woff2",
  variable: "--font-pretendard",
  display: "swap",
});

const jetbrainsMono = JetBrains_Mono({
  subsets: ["latin"],
  variable: "--font-jetbrains",
  display: "swap",
});

const fraunces = Fraunces({
  subsets: ["latin"],
  variable: "--font-fraunces",
  display: "swap",
});

export const metadata: Metadata = {
  title: "시트기반 파이썬 (Sheet Python)",
  description:
    "브라우저에서 서버 없이 실행되는 Python in Excel 스타일 워크북. 표를 붙여넣고 xl() 참조로 Python을 실행하세요.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html
      lang="ko"
      className={`${pretendard.variable} ${jetbrainsMono.variable} ${fraunces.variable}`}
    >
      <body className="antialiased">
        {children}
        {/* 부록 P.7: 오른쪽 속성 창·아래 시트 탭 줄을 가리지 않도록 위쪽 가운데 */}
        <Toaster position="top-center" />
        {/* glide-data-grid 오버레이 편집기 포털 — 없으면 셀 타이핑 편집이 조용히 실패한다 */}
        <div id="portal" />
      </body>
    </html>
  );
}
