export const panel = "min-w-0 rounded-panel border border-border bg-surface p-4 min-[421px]:p-6";
export const hint = "text-small text-muted";
export const button =
  "min-h-11 rounded-panel border border-border bg-surface px-4 py-2 font-semibold text-ink hover:bg-page disabled:cursor-not-allowed disabled:bg-page disabled:text-muted";
export const primaryButton = `${button} enabled:border-primary enabled:bg-primary enabled:text-surface enabled:hover:bg-primary-hover`;
export const textInput =
  "block w-full rounded-panel border border-border bg-page px-3 py-2 text-ink aria-invalid:border-danger disabled:cursor-not-allowed disabled:opacity-60";
export const errorMessage =
  "mt-4 whitespace-pre-wrap wrap-anywhere border-l-3 border-danger bg-danger-bg p-3 text-danger";
export const statusMessage =
  "my-3 min-h-6 text-small text-muted data-[state=success]:text-success data-[state=error]:text-danger";
