import { QuartzComponentConstructor, QuartzComponentProps } from "./types"

function RenderExcel() {
  return null
}

RenderExcel.afterDOMLoaded = `
  document.addEventListener("nav", () => {
    document.querySelectorAll('a[href$=".xlsx"]').forEach(link => {
      // Check if we already added an iframe
      if (link.nextElementSibling && link.nextElementSibling.tagName === 'DIV') return;
      
      const url = new URL(link.getAttribute('href'), window.location.href).href;
      const viewerUrl = "https://view.officeapps.live.com/op/embed.aspx?src=" + encodeURIComponent(url);
      
      const container = document.createElement('div');
      container.style.marginTop = "1rem";
      container.style.marginBottom = "2rem";

      const iframe = document.createElement('iframe');
      iframe.src = viewerUrl;
      iframe.width = "100%";
      iframe.height = "600px";
      iframe.style.border = "1px solid rgba(255, 255, 255, 0.1)";
      iframe.style.borderRadius = "8px";
      
      container.appendChild(iframe);
      link.parentNode.insertBefore(container, link.nextSibling);
      
      link.style.display = "inline-flex";
      link.style.alignItems = "center";
      link.style.gap = "0.5rem";
      link.style.padding = "0.5rem 1rem";
      link.style.backgroundColor = "var(--secondary)";
      link.style.color = "var(--darkBg)";
      link.style.borderRadius = "4px";
      link.style.textDecoration = "none";
      link.style.fontWeight = "bold";
      link.style.marginBottom = "0.5rem";
      
      // Prepend a download icon
      link.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg> Download Excel File';
    });
  });
`

export default (() => RenderExcel) satisfies QuartzComponentConstructor
