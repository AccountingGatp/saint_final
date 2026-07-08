"use client";

import { useRef, useState } from "react";
import {
  Upload,
  FileSpreadsheet,
  Loader2,
  CheckCircle2,
  Download,
  FileOutput,
  X,
} from "lucide-react";
import { Button } from "@/components/ui/button";

// const API_URL =
//   process.env.NEXT_PUBLIC_API_URL || "http://localhost:4000";

const API_URL =  "https://saint-final-api.vercel.app";

function formatSize(bytes: number) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

type StoredFile = {
  fileId: string;
  fileName: string;
  downloadUrl: string;
};

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const [result, setResult] = useState<StoredFile | null>(null);
  const [importLoading, setImportLoading] = useState(false);
  const [importError, setImportError] = useState<string | null>(null);
  const [importResult, setImportResult] = useState<StoredFile | null>(null);

  function pickFile() {
    inputRef.current?.click();
  }

  function clearImport() {
    setImportError(null);
    setImportResult(null);
  }

  function clearResult() {
    clearImport();
    setResult(null);
  }

  function selectFile(f: File | null) {
    setError(null);
    clearResult();
    setFile(f);
  }

  function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    selectFile(e.target.files?.[0] ?? null);
  }

  function onDrop(e: React.DragEvent) {
    e.preventDefault();
    setDragging(false);
    selectFile(e.dataTransfer.files?.[0] ?? null);
  }

  function clearFile() {
    selectFile(null);
    if (inputRef.current) inputRef.current.value = "";
  }

  async function handleUpload() {
    if (!file) return;
    setLoading(true);
    setError(null);
    clearResult();

    try {
      const formData = new FormData();
      formData.append("file", file);

      const res = await fetch(`${API_URL}/process`, {
        method: "POST",
        body: formData,
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data?.error || `Request failed (${res.status})`);
      }

      setResult({
        fileId: data.fileId,
        fileName: data.fileName,
        downloadUrl: data.downloadUrl,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Something went wrong.");
    } finally {
      setLoading(false);
    }
  }

  async function handleGenerateImport() {
    if (!result) return;
    setImportLoading(true);
    clearImport();

    try {
      const res = await fetch(`${API_URL}/import`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileId: result.fileId }),
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(data?.error || `Request failed (${res.status})`);
      }

      setImportResult({
        fileId: data.fileId,
        fileName: data.fileName,
        downloadUrl: data.downloadUrl,
      });
    } catch (err) {
      setImportError(
        err instanceof Error ? err.message : "Something went wrong."
      );
    } finally {
      setImportLoading(false);
    }
  }

  return (
    <main className="relative min-h-screen overflow-hidden bg-slate-50 text-slate-900">
      <div className="pointer-events-none absolute inset-0">
        <div className="absolute -top-32 left-1/2 h-96 w-96 -translate-x-1/2 rounded-full bg-indigo-300/40 blur-3xl" />
        <div className="absolute bottom-0 right-0 h-80 w-80 rounded-full bg-sky-300/40 blur-3xl" />
      </div>

      <div className="relative flex min-h-screen items-center justify-center p-4">
        <div className="w-full max-w-md">
          <div className="mb-8 text-center">
            <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-500 to-sky-400 shadow-lg shadow-indigo-500/30">
              <FileSpreadsheet className="h-7 w-7 text-white" />
            </div>
            <h1 className="bg-gradient-to-r from-indigo-600 to-sky-500 bg-clip-text text-4xl font-bold tracking-tight text-transparent">
              SAINT
            </h1>
            <p className="mt-2 text-sm text-slate-500">
              Upload an Excel file — processed and stored securely in the cloud.
            </p>
          </div>

          <div className="rounded-2xl border border-slate-200 bg-white p-6 shadow-xl shadow-slate-200/60">
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={onFileChange}
              className="hidden"
            />

            {!file ? (
              <div
                onClick={pickFile}
                onDragOver={(e) => {
                  e.preventDefault();
                  setDragging(true);
                }}
                onDragLeave={() => setDragging(false)}
                onDrop={onDrop}
                className={`flex cursor-pointer flex-col items-center justify-center gap-3 rounded-xl border-2 border-dashed px-4 py-12 text-center transition-colors ${
                  dragging
                    ? "border-indigo-400 bg-indigo-50"
                    : "border-slate-200 hover:border-indigo-400/60 hover:bg-slate-50"
                }`}
              >
                <div className="flex h-11 w-11 items-center justify-center rounded-full bg-slate-100">
                  <Upload className="h-5 w-5 text-slate-500" />
                </div>
                <div>
                  <p className="text-sm font-medium text-slate-700">
                    Drag &amp; drop your file here
                  </p>
                  <p className="mt-1 text-xs text-slate-400">
                    or click to browse · .xlsx
                  </p>
                </div>
              </div>
            ) : (
              <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-slate-50 px-4 py-3">
                <div className="flex h-10 w-10 flex-shrink-0 items-center justify-center rounded-lg bg-gradient-to-br from-indigo-500 to-sky-400">
                  <FileSpreadsheet className="h-5 w-5 text-white" />
                </div>
                <div className="min-w-0 flex-1">
                  <p className="truncate text-sm font-medium text-slate-800">
                    {file.name}
                  </p>
                  <p className="text-xs text-slate-400">
                    {formatSize(file.size)}
                  </p>
                </div>
                <button
                  onClick={clearFile}
                  disabled={loading}
                  className="flex h-8 w-8 flex-shrink-0 items-center justify-center rounded-md text-slate-400 transition-colors hover:bg-slate-200 hover:text-slate-600 disabled:opacity-50"
                  aria-label="Remove file"
                >
                  <X className="h-4 w-4" />
                </button>
              </div>
            )}

            {result ? (
              <>
                <a
                  href={result.downloadUrl}
                  download={result.fileName}
                  className="mt-4 flex h-11 w-full items-center justify-center gap-2 rounded-md bg-gradient-to-r from-emerald-500 to-green-500 text-sm font-medium text-white transition-colors hover:from-emerald-500/90 hover:to-green-500/90"
                >
                  <Download className="h-4 w-4" />
                  Download {result.fileName}
                </a>
                <div className="mt-3 flex items-center justify-center gap-2 text-sm text-emerald-600">
                  <CheckCircle2 className="h-4 w-4" />
                  Ready — click above to download.
                </div>

                <div className="mt-4 border-t border-slate-100 pt-4">
                  {importResult ? (
                    <a
                      href={importResult.downloadUrl}
                      download={importResult.fileName}
                      className="flex h-11 w-full items-center justify-center gap-2 rounded-md bg-gradient-to-r from-violet-500 to-fuchsia-500 text-sm font-medium text-white transition-colors hover:from-violet-500/90 hover:to-fuchsia-500/90"
                    >
                      <Download className="h-4 w-4" />
                      Download {importResult.fileName}
                    </a>
                  ) : (
                    <Button
                      variant="outline"
                      className="h-11 w-full border-slate-200"
                      onClick={handleGenerateImport}
                      disabled={importLoading}
                    >
                      {importLoading ? (
                        <>
                          <Loader2 className="h-4 w-4 animate-spin" />
                          Generating…
                        </>
                      ) : (
                        <>
                          <FileOutput className="h-4 w-4" />
                          Generate Import Excel
                        </>
                      )}
                    </Button>
                  )}
                  {importError && (
                    <p className="mt-3 text-center text-sm text-red-500">
                      {importError}
                    </p>
                  )}
                </div>
              </>
            ) : (
              <Button
                className="mt-4 h-11 w-full bg-gradient-to-r from-indigo-500 to-sky-500 text-white hover:from-indigo-500/90 hover:to-sky-500/90"
                onClick={handleUpload}
                disabled={!file || loading}
              >
                {loading ? (
                  <>
                    <Loader2 className="h-4 w-4 animate-spin" />
                    Processing…
                  </>
                ) : (
                  "Upload & Process"
                )}
              </Button>
            )}
            {error && (
              <p className="mt-4 text-center text-sm text-red-500">{error}</p>
            )}
          </div>

          <p className="mt-6 text-center text-xs text-slate-400">
            Files stored in Backblaze B2 · secure cloud storage
          </p>
        </div>
      </div>
    </main>
  );
}
