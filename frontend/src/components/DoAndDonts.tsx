const DOS = [
  "Upload a clear scan or digital PDF of the invoice.",
  "Keep the page straight — flat on the scanner or table.",
  "Make sure every detail is printed and readable (part no, qty, rate, amount).",
];

const DONTS = [
  "Don’t upload crooked or tilted (“cross”) pages.",
  "Don’t use blurry mobile phone photos.",
  "Don’t upload dark, cut-off, or incomplete scans.",
];

export function DoAndDonts() {
  return (
    <div className="grid w-full gap-3 sm:grid-cols-2">
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
  );
}
