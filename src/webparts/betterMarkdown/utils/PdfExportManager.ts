export class PdfExportManager {
  /**
   * Export markdown content to PDF via browser print dialog
   */
  public static async exportToPdf(tocHtml: string, mainHtml: string): Promise<void> {
    try {
      // Create a new window for printing
      const printWindow = window.open('', '_blank', 'width=800,height=600');
      if (!printWindow) {
        alert('Please allow popups to export PDF');
        return;
      }

      // Get the current styles by reading the computed styles
      const styleSheets = Array.from(document.styleSheets)
        .map(sheet => {
          try {
            return Array.from(sheet.cssRules)
              .map(rule => rule.cssText)
              .join('\n');
          } catch (e) {
            return '';
          }
        })
        .join('\n');

      // Build the print document with TOC as first page
      const printContent = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Export - Better Markdown</title>
  <style>
    ${styleSheets}

    /* Print-specific styles */
    @media print {
      @page {
        margin: 1in;
        size: letter;
      }

      body {
        margin: 0;
        padding: 0;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      }

      .toc-page {
        page-break-after: always;
        padding: 2rem;
      }

      .toc-page h2 {
        font-size: 2rem;
        margin-bottom: 2rem;
        border-bottom: 2px solid #333;
        padding-bottom: 1rem;
      }

      .content-page {
        padding: 2rem;
      }

      /* Ensure code blocks don't break across pages */
      pre, blockquote, table {
        page-break-inside: avoid;
      }

      /* Hide interactive elements */
      button, .toolbar, .toolbarItem {
        display: none !important;
      }

      /* Adjust link colors for print */
      a {
        color: #0066cc;
        text-decoration: none;
      }

      a[href^="http"]:after {
        content: " (" attr(href) ")";
        font-size: 0.8em;
        color: #666;
      }
    }

    @media screen {
      body {
        padding: 2rem;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      }

      .print-instructions {
        background: #e3f2fd;
        border: 2px solid #2196f3;
        padding: 1rem;
        margin-bottom: 2rem;
        border-radius: 4px;
      }
    }
  </style>
</head>
<body>
  <div class="print-instructions">
    <h3>📄 Export to PDF Instructions:</h3>
    <ol>
      <li>Press <strong>Ctrl+P</strong> (or Cmd+P on Mac) to open the print dialog</li>
      <li>Select "Save as PDF" as the destination</li>
      <li>Adjust settings if needed (margins, headers/footers)</li>
      <li>Click "Save"</li>
    </ol>
  </div>

  ${tocHtml ? `
  <div class="toc-page">
    <h2>📑 Table of Contents</h2>
    ${tocHtml}
  </div>
  ` : ''}

  <div class="content-page">
    ${mainHtml}
  </div>
</body>
</html>`;

      printWindow.document.write(printContent);
      printWindow.document.close();

      // Wait for content to load, then show print dialog
      printWindow.onload = () => {
        setTimeout(() => {
          printWindow.print();
        }, 500);
      };

    } catch (e) {
      console.error('PDF export error:', e);
      alert('Failed to export PDF. Please try again.');
    }
  }
}
