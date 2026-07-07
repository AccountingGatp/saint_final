import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "SAINT",
  description: "Upload an Excel file and get a processed workbook back.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
