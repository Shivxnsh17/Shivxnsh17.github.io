import { QuartzComponentConstructor, QuartzComponentProps } from "./types"

function RenderExcel() {
  return null
}

RenderExcel.afterDOMLoaded = `
  document.addEventListener("nav", () => {
    document.querySelectorAll('a[href$=".xlsx"]').forEach(link => {
      // Check if we already processed this link
      if (link.dataset.excelRendered) return;
      link.dataset.excelRendered = "true";
      
      const url = new URL(link.getAttribute('href'), window.location.href).href;
      const viewerUrl = "https://view.officeapps.live.com/op/embed.aspx?src=" + encodeURIComponent(url);
      
      // Find the block-level parent (h1-h6, p, li)
      let blockParent = link;
      while (blockParent && !['H1','H2','H3','H4','H5','H6','P','LI'].includes(blockParent.tagName)) {
          if (blockParent === document.body || blockParent.parentElement === null) break;
          blockParent = blockParent.parentElement;
      }
      
      // Insert the iframe AFTER the description text.
      // Collect siblings until we hit another header or hr.
      let insertAfterNode = blockParent;
      while (insertAfterNode.nextElementSibling) {
          let next = insertAfterNode.nextElementSibling;
          if (['H1','H2','H3','H4','H5','H6','HR'].includes(next.tagName)) {
              break;
          }
          insertAfterNode = next;
      }
      
      const container = document.createElement('div');
      container.style.marginTop = "2rem";
      container.style.marginBottom = "3rem";
      container.style.display = "flex";
      container.style.flexDirection = "column";
      container.style.gap = "1rem";

      const iframe = document.createElement('iframe');
      iframe.src = viewerUrl;
      iframe.width = "100%";
      iframe.height = "500px";
      iframe.style.border = "1px solid rgba(255, 255, 255, 0.1)";
      iframe.style.borderRadius = "8px";
      
      const downloadBtn = document.createElement('a');
      downloadBtn.href = link.getAttribute('href');
      downloadBtn.style.display = "inline-flex";
      downloadBtn.style.alignItems = "center";
      downloadBtn.style.gap = "0.5rem";
      downloadBtn.style.padding = "0.5rem 1rem";
      downloadBtn.style.backgroundColor = "var(--secondary)";
      downloadBtn.style.color = "var(--darkBg)";
      downloadBtn.style.borderRadius = "4px";
      downloadBtn.style.textDecoration = "none";
      downloadBtn.style.fontWeight = "bold";
      downloadBtn.style.fontSize = "0.875rem";
      downloadBtn.style.alignSelf = "flex-start"; // Prevents full-width stretching
      downloadBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg> Download File';
      
      container.appendChild(iframe);
      container.appendChild(downloadBtn);
      
      insertAfterNode.parentNode.insertBefore(container, insertAfterNode.nextSibling);
    });
  });
`

export default (() => RenderExcel) satisfies QuartzComponentConstructor
