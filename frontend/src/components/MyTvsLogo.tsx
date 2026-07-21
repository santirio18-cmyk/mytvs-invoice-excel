type Props = {
  className?: string;
};

/** myTVS wordmark — orange "my" + blue "TVS". */
export function MyTvsLogo({ className = "" }: Props) {
  return (
    <span
      className={`inline-flex items-baseline select-none font-[family-name:var(--font-brand)] font-extrabold tracking-tight leading-none ${className}`}
      aria-label="myTVS"
    >
      <span className="text-tvs-orange">my</span>
      <span className="text-tvs-blue">TVS</span>
    </span>
  );
}
