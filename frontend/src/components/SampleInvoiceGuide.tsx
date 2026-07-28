"use client";

const LEGEND = [
  {
    id: "1",
    label: "Supplier name",
    detail: "Seller header — sheet name.",
    tone: "blue" as const,
  },
  {
    id: "2",
    label: "Buyer / place of supply",
    detail: "Buyer + GSTIN state code.",
    tone: "blue" as const,
  },
  {
    id: "3",
    label: "Invoice no & date",
    detail: "Bill No + Date — sheet title.",
    tone: "orange" as const,
  },
  {
    id: "4",
    label: "Line items",
    detail: "Part No · Qty · Rate · Amount · HSN/SAC.",
    tone: "orange" as const,
  },
];

const toneBadge = {
  orange: "bg-tvs-orange text-white",
  blue: "bg-tvs-blue text-white",
} as const;

export function SampleInvoiceGuide() {
  return (
    <aside className="animate-rise-delay-2 flex flex-col gap-2.5 text-ink-soft md:sticky md:top-3">
      <div className="flex items-baseline justify-between gap-2">
        <p className="text-xs font-bold uppercase tracking-[0.2em] text-tvs-orange">Sample invoice</p>
        <p className="font-[family-name:var(--font-brand)] text-sm font-extrabold text-tvs-blue">
          What we extract
        </p>
      </div>

      <div className="border border-line bg-white">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img
          src="/sample-guide.jpg?v=13"
          alt="Sample invoice with boxes 1–4 highlighting supplier, bill no & date, buyer, and line items"
          width={800}
          height={632}
          className="block h-auto w-full"
          draggable={false}
        />
      </div>

      <ol className="grid grid-cols-2 gap-x-3 gap-y-2 border border-line bg-white px-3 py-2.5">
        {LEGEND.map((c) => (
          <li key={c.id} className="flex min-w-0 items-start gap-2">
            <span
              className={`mt-px flex h-5 w-5 shrink-0 items-center justify-center text-[11px] font-bold leading-none ${toneBadge[c.tone]}`}
            >
              {c.id}
            </span>
            <div className="min-w-0">
              <p className="text-xs font-bold leading-tight text-ink">{c.label}</p>
              <p className="mt-0.5 text-[11px] leading-snug text-ink-soft/90">{c.detail}</p>
            </div>
          </li>
        ))}
      </ol>
    </aside>
  );
}
