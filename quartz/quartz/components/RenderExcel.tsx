import { QuartzComponentConstructor, QuartzComponentProps } from "./types"

function RenderExcel() {
  return null
}

RenderExcel.afterDOMLoaded = `
  document.addEventListener("nav", () => {
    // Check if this page is marked view-only (has the comment in the source)
    const isViewOnly = document.body.innerHTML.includes('view-only: no download link intentional');

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
      // Prevent right-click save on the iframe
      iframe.addEventListener('contextmenu', e => e.preventDefault());
      
      container.appendChild(iframe);

      // Only show download button if NOT view-only
      if (!isViewOnly) {
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
        downloadBtn.style.alignSelf = "flex-start";
        downloadBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg> Download File';
        container.appendChild(downloadBtn);
      } else {
        // View-only badge
        const badge = document.createElement('div');
        badge.style.display = "inline-flex";
        badge.style.alignItems = "center";
        badge.style.gap = "0.4rem";
        badge.style.padding = "0.4rem 0.8rem";
        badge.style.border = "1px solid rgba(255,255,255,0.15)";
        badge.style.borderRadius = "4px";
        badge.style.fontSize = "0.75rem";
        badge.style.color = "rgba(255,255,255,0.4)";
        badge.style.alignSelf = "flex-start";
        badge.style.userSelect = "none";
        badge.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path></svg> View Only — © Shivansh Chandra';
        container.appendChild(badge);
      }
      
      insertAfterNode.parentNode.insertBefore(container, insertAfterNode.nextSibling);
    });
  });
`

export default (() => RenderExcel) satisfies QuartzComponentConstructor
