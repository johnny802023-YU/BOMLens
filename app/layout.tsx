import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import { headers } from "next/headers";
import "./globals.css";

const geistSans = Geist({ variable: "--font-geist-sans", subsets: ["latin"] });
const geistMono = Geist_Mono({ variable: "--font-geist-mono", subsets: ["latin"] });

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host = requestHeaders.get("x-forwarded-host") ?? requestHeaders.get("host") ?? "localhost:3000";
  const protocol = requestHeaders.get("x-forwarded-proto") ?? (host.includes("localhost") ? "http" : "https");
  const imageUrl = `${protocol}://${host}/og.png`;
  return {
    title: "BOMLens｜BOM 與線路圖版本比對",
    description: "快速找出舊版與新版 BOM、線路圖的新增、刪除和內容變更。",
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: {
      title: "BOMLens｜BOM 與線路圖版本比對",
      description: "工程變更，一眼看清楚。",
      type: "website",
      images: [{ url: imageUrl, width: 1731, height: 909, alt: "BOMLens BOM 與線路圖版本比對" }],
    },
    twitter: { card: "summary_large_image", title: "BOMLens", description: "BOM 與線路圖版本比對", images: [imageUrl] },
  };
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-Hant"><body className={`${geistSans.variable} ${geistMono.variable}`}>{children}</body></html>;
}
