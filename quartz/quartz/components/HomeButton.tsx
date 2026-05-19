import { QuartzComponent, QuartzComponentConstructor, QuartzComponentProps } from "./types"

const HomeButton: QuartzComponent = (_props: QuartzComponentProps) => {
  return (
    <a href="/" id="portfolio-home-btn" aria-label="Back to Portfolio Homepage">
      {/* House icon */}
      <svg
        xmlns="http://www.w3.org/2000/svg"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
      >
        <path d="M3 9.5L12 3l9 6.5V20a1 1 0 01-1 1H4a1 1 0 01-1-1V9.5z" />
        <path d="M9 21V12h6v9" />
      </svg>
      <span>Portfolio</span>
    </a>
  )
}

HomeButton.css = `
#portfolio-home-btn {
  position: fixed;
  top: 1rem;
  left: 1rem;
  z-index: 9999;
  display: inline-flex;
  align-items: center;
  gap: 0.4rem;
  padding: 0.45rem 0.9rem 0.45rem 0.7rem;
  font-size: 0.75rem;
  font-weight: 700;
  letter-spacing: 0.06em;
  text-transform: uppercase;
  color: var(--secondary);
  background: rgba(11, 15, 25, 0.85);
  border: 1px solid rgba(56, 189, 248, 0.35);
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
  clip-path: polygon(8px 0%, 100% 0%, 100% calc(100% - 8px), calc(100% - 8px) 100%, 0% 100%, 0% 8px);
  text-decoration: none !important;
  transition: background 0.2s ease, border-color 0.2s ease, transform 0.2s ease, box-shadow 0.2s ease;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.4);
  white-space: nowrap;
}

#portfolio-home-btn:hover {
  background: rgba(56, 189, 248, 0.15);
  border-color: rgba(56, 189, 248, 0.65);
  transform: translateY(-1px);
  box-shadow: 0 6px 20px rgba(56, 189, 248, 0.2);
  text-decoration: none !important;
}

#portfolio-home-btn svg {
  width: 0.95rem;
  height: 0.95rem;
  flex-shrink: 0;
  stroke: var(--secondary);
}

@media (max-width: 768px) {
  #portfolio-home-btn span {
    display: none;
  }
  #portfolio-home-btn {
    padding: 0.5rem;
    top: 0.6rem;
    left: 0.6rem;
  }
  #portfolio-home-btn svg {
    width: 1.1rem;
    height: 1.1rem;
  }
}
`

export default (() => HomeButton) satisfies QuartzComponentConstructor
