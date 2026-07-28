"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { MyTvsLogo } from "../components/MyTvsLogo";
import { SampleInvoiceGuide } from "../components/SampleInvoiceGuide";
import { DoAndDonts } from "../components/DoAndDonts";

const MAX_FILES = 10;
/**
 * API base URL:
 * - Local: http://127.0.0.1:8000
 * - Vercel: same origin (/api/...) — rewritten to Railway in vercel.json
 * - Override anytime with NEXT_PUBLIC_API_URL
 */
const API_BASE =
  process.env.NEXT_PUBLIC_API_URL ||
  (typeof window !== "undefined" && window.location.hostname === "localhost"
    ? "http://127.0.0.1:8000"
    : "");

type Status = "idle" | "converting" | "done" | "error";

type LineItem = {
  part_number?: string;
  description?: string;
  hsn_sac?: string;
  qty?: string;
  rate?: string;
  amount?: string;
};

type InvoiceResult = {
  invoice_number: string;
  supplier_name: string;
  date: string;
  place_of_supply: string;
  filename: string;
  item_count: number;
  line_items: LineItem[];
};

function formatBytes(n: number) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  return url;
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState<Status>("idle");
  const [message, setMessage] = useState("");
  const [results, setResults] = useState<InvoiceResult[]>([]);
  const [csvUrl, setCsvUrl] = useState<string | null>(null);
  const [zipUrl, setZipUrl] = useState<string | null>(null);
  const [outputDir, setOutputDir] = useState("");

  const canAddMore = files.length < MAX_FILES;

  const isAllowed = (f: File) => {
    const n = f.name.toLowerCase();
    return (
      f.type === "application/pdf" ||
      f.type.startsWith("image/") ||
      n.endsWith(".pdf") ||
      n.endsWith(".png") ||
      n.endsWith(".jpg") ||
      n.endsWith(".jpeg") ||
      n.endsWith(".webp")
    );
  };

  const addFiles = useCallback(
    (incoming: FileList | File[]) => {
      const list = Array.from(incoming).filter(isAllowed);
      setFiles((prev) => {
        const room = MAX_FILES - prev.length;
        if (room <= 0) return prev;
        const next = [...prev];
        for (const f of list.slice(0, room)) {
          if (!next.some((x) => x.name === f.name && x.size === f.size)) {
            next.push(f);
          }
        }
        return next;
      });
      setStatus("idle");
      setMessage("");
      setResults([]);
      if (csvUrl) URL.revokeObjectURL(csvUrl);
      if (zipUrl) URL.revokeObjectURL(zipUrl);
      setCsvUrl(null);
      setZipUrl(null);
      setOutputDir("");
    },
    [csvUrl, zipUrl],
  );

  const removeFile = (index: number) => {
    setFiles((prev) => prev.filter((_, i) => i !== index));
    setStatus("idle");
    setMessage("");
  };

  const clearAll = () => {
    setFiles([]);
    setStatus("idle");
    setMessage("");
    setResults([]);
    if (csvUrl) URL.revokeObjectURL(csvUrl);
    if (zipUrl) URL.revokeObjectURL(zipUrl);
    setCsvUrl(null);
    setZipUrl(null);
    setOutputDir("");
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    if (e.dataTransfer.files?.length) addFiles(e.dataTransfer.files);
  };

  const convert = async () => {
    if (!files.length) {
      setStatus("error");
      setMessage("Add at least one invoice.");
      return;
    }

    setStatus("converting");
    setMessage("Extracting invoice data…");
    setResults([]);
    if (csvUrl) URL.revokeObjectURL(csvUrl);
    if (zipUrl) URL.revokeObjectURL(zipUrl);
    setCsvUrl(null);
    setZipUrl(null);

    try {
      const form = new FormData();
      files.forEach((f) => form.append("files", f));

      const parseRes = await fetch(`${API_BASE}/api/parse`, { method: "POST", body: form });
      if (!parseRes.ok) {
        let detail = "Extraction failed.";
        try {
          const data = await parseRes.json();
          detail = data.detail || detail;
        } catch {
          /* ignore */
        }
        throw new Error(typeof detail === "string" ? detail : JSON.stringify(detail));
      }
      const parsed = await parseRes.json();
      const invoices: InvoiceResult[] = parsed.invoices || [];
      setResults(invoices);
      setOutputDir(parsed.output_dir || parsed.latest_csv || "");

      const totalItems = invoices.reduce((n, inv) => n + (inv.item_count || 0), 0);

      const formCsv = new FormData();
      files.forEach((f) => formCsv.append("files", f));
      const csvRes = await fetch(`${API_BASE}/api/convert-csv`, { method: "POST", body: formCsv });
      if (!csvRes.ok) throw new Error("CSV download failed.");
      const csvBlob = await csvRes.blob();
      if (csvBlob.size < 20) throw new Error("CSV came back empty — is backend running?");
      const csvName =
        (csvRes.headers.get("Content-Disposition") || "").match(/filename="?([^"]+)"?/)?.[1] ||
        "invoices.csv";
      const nextCsvUrl = triggerDownload(csvBlob, csvName);
      setCsvUrl(nextCsvUrl);

      const formZip = new FormData();
      files.forEach((f) => formZip.append("files", f));
      const zipRes = await fetch(`${API_BASE}/api/convert`, { method: "POST", body: formZip });
      if (zipRes.ok) {
        const zipBlob = await zipRes.blob();
        const zipName =
          (zipRes.headers.get("Content-Disposition") || "").match(/filename="?([^"]+)"?/)?.[1] ||
          "invoices.zip";
        const nextZipUrl = triggerDownload(zipBlob, zipName);
        setZipUrl(nextZipUrl);
      }

      setStatus(totalItems ? "done" : "error");
      setMessage(
        totalItems
          ? `Success: ${invoices.length} invoice(s), ${totalItems} line item(s). CSV downloaded (opens in Numbers). Also saved to: ${parsed.latest_csv || "output/latest.csv"}`
          : `Headers found but 0 line items. Open preview below. File saved to: ${parsed.latest_csv || "output/latest.csv"}`,
      );
    } catch (err) {
      setStatus("error");
      setMessage(err instanceof Error ? err.message : "Something went wrong.");
    }
  };

  const hint = useMemo(() => {
    if (!files.length) return "Drop PDF or photo invoices — up to 10 files";
    return `${files.length} of ${MAX_FILES} selected`;
  }, [files.length]);

  return (
    <main className="relative flex flex-1 flex-col px-5 pb-12 pt-6 sm:px-8 lg:px-12">
      <div className="mx-auto flex w-full max-w-6xl flex-1 flex-col">
        <header className="animate-rise mb-6 sm:mb-8">
          <div className="animate-brand mb-3 sm:mb-4">
            <MyTvsLogo className="text-[2.5rem] sm:text-5xl lg:text-6xl" />
          </div>
          <p className="max-w-xl text-sm font-medium leading-relaxed text-ink-soft sm:text-base">
            Invoice PDFs &amp; photos → Excel. Part number, HSN/SAC, qty, rate, and amount on every line.
          </p>
        </header>

        <section className="animate-rise-delay grid flex-1 gap-5 md:grid-cols-2 md:items-start md:gap-6">
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            className={`relative overflow-hidden border border-line bg-white/75 p-4 backdrop-blur-sm transition sm:p-6 ${
              dragging ? "drop-active border-tvs-orange bg-tvs-orange/5" : ""
            }`}
          >
            <div className="relative flex flex-col items-start gap-4">
              <div className="flex w-full items-end justify-between gap-4">
                <div>
                  <p className="text-xs font-bold uppercase tracking-[0.2em] text-tvs-orange">Upload</p>
                  <p className="mt-1 text-sm text-ink-soft/85">{hint}</p>
                </div>
                {files.length > 0 && (
                  <button
                    type="button"
                    onClick={clearAll}
                    className="text-sm text-ink-soft/70 underline-offset-4 hover:text-tvs-blue hover:underline"
                  >
                    Clear
                  </button>
                )}
              </div>

              <button
                type="button"
                disabled={!canAddMore}
                onClick={() => inputRef.current?.click()}
                className="group flex w-full flex-col items-center justify-center gap-2 border border-dashed border-tvs-blue/25 bg-paper/70 px-5 py-8 text-center transition hover:border-tvs-orange hover:bg-tvs-orange/5 disabled:cursor-not-allowed disabled:opacity-50"
              >
                <span className="flex h-10 w-10 items-center justify-center rounded-full bg-tvs-blue text-white transition group-hover:scale-105 group-hover:bg-tvs-orange">
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden>
                    <path
                      d="M12 16V4M12 4l-4 4M12 4l4 4M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2"
                      stroke="currentColor"
                      strokeWidth="1.8"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    />
                  </svg>
                </span>
                <span className="text-sm font-semibold text-ink sm:text-base">
                  {canAddMore ? "Choose invoices" : "Maximum 10 files reached"}
                </span>
                <span className="text-xs text-ink-soft/70 sm:text-sm">PDF or photos (PNG/JPG) · no login</span>
              </button>

              <DoAndDonts />

              <input
                ref={inputRef}
                type="file"
                accept="application/pdf,.pdf,image/png,image/jpeg,image/webp,.png,.jpg,.jpeg,.webp"
                multiple
                className="hidden"
                onChange={(e) => {
                  if (e.target.files) addFiles(e.target.files);
                  e.target.value = "";
                }}
              />

              {files.length > 0 && (
                <ul className="w-full divide-y divide-line border border-line bg-white">
                  {files.map((f, i) => (
                    <li key={`${f.name}-${f.size}-${i}`} className="flex items-center justify-between gap-3 px-4 py-3">
                      <div className="min-w-0">
                        <p className="truncate text-sm font-semibold text-ink">{f.name}</p>
                        <p className="text-xs text-ink-soft/70">{formatBytes(f.size)}</p>
                      </div>
                      <button
                        type="button"
                        onClick={() => removeFile(i)}
                        className="shrink-0 text-xs font-bold uppercase tracking-wide text-warn hover:opacity-80"
                      >
                        Remove
                      </button>
                    </li>
                  ))}
                </ul>
              )}

              <div className="flex flex-wrap gap-3">
                <button
                  type="button"
                  onClick={convert}
                  disabled={!files.length || status === "converting"}
                  className="inline-flex items-center justify-center gap-2 bg-tvs-blue px-6 py-3.5 text-sm font-bold tracking-wide text-white transition hover:bg-tvs-blue-deep disabled:cursor-not-allowed disabled:opacity-45"
                >
                  {status === "converting" ? (
                    <>
                      <span className="spinner" />
                      Converting…
                    </>
                  ) : (
                    "Generate Excel / CSV"
                  )}
                </button>

                {csvUrl && (
                  <a
                    href={csvUrl}
                    download="invoices.csv"
                    className="inline-flex items-center justify-center bg-tvs-orange px-6 py-3.5 text-sm font-bold tracking-wide text-white transition hover:bg-tvs-orange-deep"
                  >
                    Download CSV again
                  </a>
                )}
                {zipUrl && (
                  <a
                    href={zipUrl}
                    download="invoices.zip"
                    className="inline-flex items-center justify-center border border-tvs-blue px-6 py-3.5 text-sm font-bold tracking-wide text-tvs-blue transition hover:bg-tvs-blue hover:text-white"
                  >
                    Download ZIP (CSV+XLSX)
                  </a>
                )}
              </div>

              {outputDir && (
                <p className="break-all text-xs text-ink-soft/80">
                  Also saved on disk: <code className="text-ink">{outputDir}</code>
                </p>
              )}

              {message && (
                <p
                  className={`text-sm leading-relaxed ${
                    status === "error"
                      ? "text-warn"
                      : status === "done"
                        ? "text-tvs-blue"
                        : "text-ink-soft"
                  }`}
                >
                  {message}
                </p>
              )}
            </div>
          </div>

          <SampleInvoiceGuide />
        </section>

        {results.length > 0 && (
          <section className="animate-rise mt-10 space-y-8 border-t border-line pt-10">
            <p className="text-xs font-bold uppercase tracking-[0.2em] text-tvs-orange">Extracted preview</p>
            {results.map((inv) => (
              <div key={`${inv.invoice_number}-${inv.filename}`} className="overflow-x-auto border border-line bg-white">
                <div className="border-b border-line px-4 py-3 text-sm text-ink">
                  <strong className="text-tvs-blue">{inv.invoice_number}</strong>
                  {" · "}
                  {inv.supplier_name}
                  {" · "}
                  {inv.date}
                  {" · "}
                  {inv.place_of_supply}
                  {" · "}
                  {inv.item_count} items
                </div>
                <table className="w-full min-w-[640px] text-left text-sm">
                  <thead className="bg-tvs-blue text-white">
                    <tr>
                      <th className="px-3 py-2.5 font-semibold">Part Number</th>
                      <th className="px-3 py-2.5 font-semibold">Description</th>
                      <th className="px-3 py-2.5 font-semibold">Qty</th>
                      <th className="px-3 py-2.5 font-semibold">Rate</th>
                      <th className="px-3 py-2.5 font-semibold">Amount</th>
                      <th className="px-3 py-2.5 font-semibold">HSN/SAC</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(inv.line_items || []).slice(0, 30).map((it, idx) => (
                      <tr key={idx} className={idx % 2 ? "bg-paper/80" : ""}>
                        <td className="px-3 py-2 text-ink">{it.part_number}</td>
                        <td className="px-3 py-2 text-ink">{it.description}</td>
                        <td className="px-3 py-2 text-ink">{it.qty}</td>
                        <td className="px-3 py-2 text-ink">{it.rate}</td>
                        <td className="px-3 py-2 text-ink">{it.amount}</td>
                        <td className="px-3 py-2 text-ink">{it.hsn_sac}</td>
                      </tr>
                    ))}
                    {!inv.line_items?.length && (
                      <tr>
                        <td colSpan={6} className="px-3 py-4 text-warn">
                          No line items detected for this file.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
                {inv.item_count > 30 && (
                  <p className="px-4 py-2 text-xs text-ink-soft/70">
                    Showing first 30 of {inv.item_count} rows — full data is in the Excel file.
                  </p>
                )}
              </div>
            ))}
          </section>
        )}
      </div>
    </main>
  );
}
