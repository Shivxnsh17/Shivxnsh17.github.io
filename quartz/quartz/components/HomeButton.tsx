import { QuartzComponent, QuartzComponentConstructor, QuartzComponentProps } from "./types"
import { classNames } from "../util/lang"

const HomeButton: QuartzComponent = ({ displayClass }: QuartzComponentProps) => {
  return (
    <a
      href="/"
      class={classNames(displayClass, "home-button")}
      aria-label="Back to Portfolio Homepage"
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="square"
        stroke-linejoin="miter"
        class="home-button-icon"
      >
        <path d="M3 12L12 3l9 9" />
        <path d="M9 21V12h6v9" />
        <path d="M3 12v9h18V12" />
      </svg>
      <span>Portfolio</span>
    </a>
  )
}

HomeButton.css = `
.home-button {
  display: inline-flex;
  align-items: center;
  gap: 0.45rem;
  padding: 0.4rem 0.85rem 0.4rem 0.65rem;
  font-size: 0.8rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--secondary);
  background: rgba(56, 189, 248, 0.08);
  border: 1px solid rgba(56, 189, 248, 0.25);
  border-radius: 0;
  clip-path: polygon(8px 0%, 100% 0%, 100% calc(100% - 8px), calc(100% - 8px) 100%, 0% 100%, 0% 8px);
  text-decoration: none;
  transition: background 0.2s ease, color 0.2s ease, border-color 0.2s ease, transform 0.2s ease;
  white-space: nowrap;
  margin-bottom: 0.5rem;
}

.home-button:hover {
  background: rgba(56, 189, 248, 0.18);
  border-color: rgba(56, 189, 248, 0.55);
  color: var(--secondary);
  transform: translateX(-2px);
  text-decoration: none;
}

.home-button-icon {
  width: 1rem;
  height: 1rem;
  flex-shrink: 0;
}
`

export default (() => HomeButton) satisfies QuartzComponentConstructor
