const EXAMPLES = [
  {
    src: "/guides/good-clear-pdf.jpg",
    badge: "Works well",
    badgeClass: "bg-tvs-blue text-white",
    title: "Clear PDF or flat scan",
    detail: "Digital invoice or straight scan — part no, qty, rate, amount readable.",
  },
  {
    src: "/guides/weak-screenshot.jpg",
    badge: "Often fails",
    badgeClass: "bg-warn text-white",
    title: "Email / screen capture",
    detail: "UI chrome, low resolution, or cropped table — line items often missed.",
  },
  {
    src: "/guides/weak-phone-photo.jpg",
    badge: "Often fails",
    badgeClass: "bg-warn text-white",
    title: "Blurry WhatsApp photo",
    detail: "Tilted, dark, or compressed phone shots — OCR struggles on rates & HSN.",
  },
] as const;

const DOS = [
  "Upload a clear scan or digital PDF of the invoice.",
  "Keep the page straight — flat on the scanner or table.",
  "Make sure every detail is printed and readable (part no, qty, rate, amount).",
];

const DONTS = [
  "Don’t upload crooked or tilted (“cross”) pages.",
  "Don’t use blurry mobile phone photos.",
  "Don’t upload email screenshots or dark, cut-off scans.",
];

export function DoAndDonts() {
  return (
    <div className="flex w-full flex-col gap-3">
      <div className="grid gap-3 sm:grid-cols-2">
        <div className="border border-tvs-blue/25 bg-tvs-blue/[0.04] p-3">
          <p className="text-xs font-bold uppercase tracking-[0.16em] text-tvs-blue">Do’s</p>
          <ul className="mt-2 space-y-1.5">
            {DOS.map((item) => (
              <li key={item} className="flex gap-2 text-[11px] leading-snug text-ink sm:text-xs">
                <span className="mt-0.5 shrink-0 font-bold text-tvs-blue" aria-hidden>
                  ✓
                </span>
                <span>{item}</span>
              </li>
            ))}
          </ul>
        </div>

        <div className="border border-warn/30 bg-warn/[0.04] p-3">
          <p className="text-xs font-bold uppercase tracking-[0.16em] text-warn">Don’ts</p>
          <ul className="mt-2 space-y-1.5">
            {DONTS.map((item) => (
              <li key={item} className="flex gap-2 text-[11px] leading-snug text-ink sm:text-xs">
                <span className="mt-0.5 shrink-0 font-bold text-warn" aria-hidden>
                  ✕
                </span>
                <span>{item}</span>
              </li>
            ))}
          </ul>
        </div>
      </div>

      <div>
        <p className="text-xs font-bold uppercase tracking-[0.16em] text-ink-soft">
          Upload quality — what still struggles
        </p>
        <p className="mt-1 text-[11px] leading-snug text-ink-soft/85 sm:text-xs">
          Clear PDFs work best. These weak cases often return missing lines or wrong qty/MRP.
        </p>
        <div className="mt-2.5 grid gap-2.5 sm:grid-cols-3">
          {EXAMPLES.map((ex) => (
            <figure
              key={ex.src}
              className="overflow-hidden border border-line bg-white"
            >
              <div className="relative aspect-[4/3] bg-paper-deep">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img
                  src={ex.src}
                  alt={ex.title}
                  className="h-full w-full object-cover object-top"
                  draggable={false}
                />
                <span
                  className={`absolute left-2 top-2 px-1.5 py-0.5 text-[10px] font-bold uppercase tracking-wide ${ex.badgeClass}`}
                >
                  {ex.badge}
                </span>
              </div>
              <figcaption className="space-y-0.5 px-2.5 py-2">
                <p className="text-xs font-bold text-ink">{ex.title}</p>
                <p className="text-[11px] leading-snug text-ink-soft/90">{ex.detail}</p>
              </figcaption>
            </figure>
          ))}
        </div>
      </div>
    </div>
  );
}
