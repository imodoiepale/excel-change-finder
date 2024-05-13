import type { Metadata, Viewport } from "next";
import { Inter } from "next/font/google";
import "./globals.css";

const inter = Inter({ subsets: ["latin"] });

export const metadata: Metadata = {
  applicationName: "EXCEL CHANGE FINDER",
  title: "EXCEL CHANGE FINDER",
  description: "SEE THE CHANGES DONE ON AN EXCEL FILE",
  manifest: "/manifest.json",
  openGraph: {
    type: "website",
    siteName: "EXCEL CHANGE FINDER",
    title: {
      default: "EXCEL CHANGE FINDER",
      template: "EXCEL CHANGE FINDER",
    },
    description: "SEE THE CHANGES DONE ON AN EXCEL FILE",
  },
};

export const viewport: Viewport = {
  themeColor: "#111827",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className={inter.className}>{children}</body>
    </html>
  );
}
