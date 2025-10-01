import { MonacoLoader } from './MonacoLoader';
import { MarkdownProcessor } from './MarkdownProcessor';
import { MermaidRenderer } from './MermaidRenderer';
import { CodeBlockEnhancer } from './CodeBlockEnhancer';
import { PropertyPaneDetector } from './PropertyPaneDetector';

export interface IEditModeManagerOptions {
  styles: any;
  properties: any;
  markdownProcessor: MarkdownProcessor;
  mermaidRenderer: MermaidRenderer;
  codeBlockEnhancer: CodeBlockEnhancer;
  propertyPaneDetector: PropertyPaneDetector;
  onPropertyChange: (property: string, value: any) => void;
  onPropertyPaneRefresh: () => void;
}

export class EditModeManager {
  private options: IEditModeManagerOptions;
  private monacoEditor: any = null;
  private editorContent: string = '';
  private isEditorCollapsed: boolean = false;
  private isSyncingFromEditor: boolean = false;
  private isSyncingFromPreview: boolean = false;

  constructor(options: IEditModeManagerOptions) {
    this.options = options;
  }

  public async renderEditMode(domElement: HTMLElement): Promise<void> {
    // Initialize editor content with current property value
    this.editorContent = this.options.properties.markdownContent || '# Wiki.js Style Markdown\n\nStart editing to see the preview...';
    
    // Render initial preview with TOC
    const initialHtml = this.options.markdownProcessor.render(this.editorContent);
    const { tocHtml, mainHtml } = this.options.markdownProcessor.extractTOC(initialHtml);
    
    // Create split-screen editor layout
    domElement.innerHTML = `<div class="${this.options.styles.betterMarkdown} ${this.options.styles.editorMode} ${this.options.styles[this.options.properties.theme] || ''}">
      <div class="${this.options.styles.editorLayout}">
        <div class="${this.options.styles.editorPane}">
          <div class="${this.options.styles.editorHeader}">
            <h3>Markdown Editor</h3>
            <div class="${this.options.styles.editorActions}">
              <button id="importMdFile" class="${this.options.styles.editorActionButton}" title="Import .md file">📁 Import</button>
              <button id="toggleEditor" class="${this.options.styles.editorActionButton}" title="Collapse/Expand editor">◀️ Collapse</button>
              <button id="saveContent" class="${this.options.styles.saveButton}">💾 Save</button>
            </div>
          </div>
          <div id="monacoContainer" class="${this.options.styles.monacoContainer}"></div>
        </div>
        <div class="${this.options.styles.resizer}"></div>
        <div class="${this.options.styles.previewPane}">
          <div class="${this.options.styles.previewHeader}">
            <h3>Live Preview</h3>
            <button id="exportPdfEdit" class="${this.options.styles.exportButton}" title="Export as PDF">📄 Export PDF</button>
          </div>
          <div class="${this.options.styles.previewContainer}">
            <div id="previewContent" class="${this.options.styles.previewContent}">
              ${mainHtml}
            </div>
          </div>
        </div>
        ${tocHtml ? `<aside class="${this.options.styles.tocSidebar}">${tocHtml}</aside>` : ''}
      </div>
    </div>`;

    // Add export PDF button functionality
    const exportPdfButton = domElement.querySelector('#exportPdfEdit') as HTMLButtonElement;
    if (exportPdfButton) {
      exportPdfButton.addEventListener('click', () => {
        void this.exportToPdf(tocHtml, mainHtml);
      });
    }

    // Add editor functionality
    await this.initializeMonacoEditor(domElement);

    // Enhance the initial content
    await this.enhanceInitialContent(domElement);
  }

  private async enhanceInitialContent(domElement: HTMLElement): Promise<void> {
    console.log('🎨 EditMode: Enhancing initial content...');

    // Enhancement target should be the entire editor layout
    const editorLayout = domElement.querySelector(`.${this.options.styles.editorLayout}`) as HTMLElement;

    if (editorLayout) {
      // Add copy button functionality to code blocks
      this.options.codeBlockEnhancer.addCopyButtonFunctionality(editorLayout);

      // Update TOC position - initial positioning without animation
      this.updateTOCPosition(domElement, this.options.propertyPaneDetector.isPropertyPaneOpen(), true);

      // Add TOC click handlers for scroll sync
      this.addTOCClickHandlers(domElement);

      // Render Mermaid diagrams for initial content
      if (this.options.properties.enableMermaid) {
        console.log('🎨 EditMode: Rendering initial Mermaid diagrams...');

        setTimeout(async () => {
          try {
            await this.options.mermaidRenderer.renderDiagrams(editorLayout, this.options.styles.mermaidError);
            console.log('🎨 EditMode: Initial Mermaid rendering completed');
          } catch (e) {
            console.error('🎨 EditMode: Initial Mermaid rendering error:', e);
          }
        }, 100);
      }
    }
  }

  private addTOCClickHandlers(domElement: HTMLElement): void {
    const tocSidebar = domElement.querySelector(`.${this.options.styles.tocSidebar}`) as HTMLElement;
    if (!tocSidebar) return;

    const tocLinks = tocSidebar.querySelectorAll('a[href^="#"]');
    tocLinks.forEach(link => {
      link.addEventListener('click', (e) => {
        // Let the default scroll behavior happen
        // Then sync the editor position after a short delay
        setTimeout(() => {
          const previewContent = domElement.querySelector('#previewContent') as HTMLElement;
          if (previewContent) {
            this.syncEditorToPreview(previewContent);
          }
        }, 100);
      });
    });
  }

  private async initializeMonacoEditor(domElement: HTMLElement): Promise<void> {
    const container = domElement.querySelector('#monacoContainer') as HTMLElement;
    const previewContent = domElement.querySelector('#previewContent') as HTMLElement;
    const saveButton = domElement.querySelector('#saveContent') as HTMLButtonElement;
    const toggleButton = domElement.querySelector('#toggleEditor') as HTMLButtonElement;
    const importButton = domElement.querySelector('#importMdFile') as HTMLButtonElement;
    const resizer = domElement.querySelector(`.${this.options.styles.resizer}`) as HTMLElement;

    if (!container || !previewContent || !saveButton || !toggleButton || !importButton) return;

    try {
      console.log('📝 EditMode: Initializing Monaco Editor...');
      
      // Create Monaco Editor
      const editorOptions = {
        value: this.editorContent,
        language: 'markdown',
        theme: this.options.properties.theme === 'dark' ? 'vs-dark' : 'vs-light',
        automaticLayout: true,
        wordWrap: 'on',
        lineNumbers: 'on',
        minimap: { enabled: false },
        scrollBeyondLastLine: false,
        fontFamily: 'Monaco, Consolas, "Courier New", monospace',
        fontSize: 14,
        lineHeight: 1.6
      };

      this.monacoEditor = await MonacoLoader.createEditor(container, editorOptions);
      console.log('✅ EditMode: Monaco Editor created successfully');

      // Real-time preview updates
      let updateTimeout: number;
      this.monacoEditor.onDidChangeModelContent(() => {
        clearTimeout(updateTimeout);
        updateTimeout = window.setTimeout(() => {
          this.editorContent = this.monacoEditor.getValue();
          this.updatePreview(domElement, previewContent);
        }, 300);
      });

      // Scroll sync: preview follows editor cursor
      this.monacoEditor.onDidChangeCursorPosition(() => {
        this.syncPreviewToEditor(previewContent);
      });

      // Scroll sync: preview follows editor scroll
      this.monacoEditor.onDidScrollChange(() => {
        this.syncPreviewToEditor(previewContent);
      });

      // Bidirectional scroll sync: editor follows preview scroll
      previewContent.addEventListener('scroll', () => {
        this.syncEditorToPreview(previewContent);
      });

      // Save functionality
      saveButton.addEventListener('click', () => {
        const content = this.monacoEditor.getValue();
        this.options.onPropertyChange('markdownContent', content);
        this.options.onPropertyPaneRefresh();
        const originalText = saveButton.textContent;
        saveButton.textContent = '✅ Saved!';
        setTimeout(() => {
          saveButton.textContent = originalText;
        }, 2000);
      });

      // Toggle editor collapse/expand
      toggleButton.addEventListener('click', () => {
        this.toggleEditorCollapse(domElement, toggleButton);
      });

      // Import .md file
      importButton.addEventListener('click', () => {
        this.importMarkdownFile();
      });

      // Keyboard shortcuts
      const monaco = (window as any).monaco;
      if (monaco && monaco.KeyMod && monaco.KeyCode) {
        this.monacoEditor.addCommand(monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS, () => {
          saveButton.click();
        });
      }

      // Resizable panes
      if (resizer) {
        this.initializeResizer(domElement, resizer);
      }

    } catch (error) {
      console.error('❌ EditMode: Failed to initialize Monaco Editor:', error);
      // Fallback to textarea if Monaco fails
      this.renderFallbackEditor(container, previewContent, saveButton, resizer, domElement);
    }
  }

  private syncPreviewToEditor(previewElement: HTMLElement): void {
    if (!this.monacoEditor || this.isSyncingFromPreview) return;

    try {
      this.isSyncingFromEditor = true;

      // Get cursor position in editor
      const position = this.monacoEditor.getPosition();
      if (!position) return;

      const currentLine = position.lineNumber;
      const totalLines = this.monacoEditor.getModel().getLineCount();

      // Calculate scroll percentage based on cursor position
      const scrollPercentage = currentLine / totalLines;

      // previewElement is #previewContent which has overflow-y: auto
      // Use its scrollHeight which represents the full content height
      const maxScroll = previewElement.scrollHeight - previewElement.clientHeight;
      const targetScroll = maxScroll * scrollPercentage;

      // Smooth scroll to position
      previewElement.scrollTo({
        top: targetScroll,
        behavior: 'smooth'
      });

      setTimeout(() => {
        this.isSyncingFromEditor = false;
      }, 100);
    } catch (e) {
      console.error('EditMode: Scroll sync error:', e);
      this.isSyncingFromEditor = false;
    }
  }

  private syncEditorToPreview(previewElement: HTMLElement): void {
    if (!this.monacoEditor || this.isSyncingFromEditor) return;

    try {
      this.isSyncingFromPreview = true;

      // Calculate scroll percentage from preview
      const scrollTop = previewElement.scrollTop;
      const maxScroll = previewElement.scrollHeight - previewElement.clientHeight;
      const scrollPercentage = maxScroll > 0 ? scrollTop / maxScroll : 0;

      // Calculate target line in editor
      const totalLines = this.monacoEditor.getModel().getLineCount();
      const targetLine = Math.max(1, Math.round(scrollPercentage * totalLines));

      // Scroll editor to target line
      this.monacoEditor.revealLineInCenter(targetLine);

      setTimeout(() => {
        this.isSyncingFromPreview = false;
      }, 100);
    } catch (e) {
      console.error('EditMode: Reverse scroll sync error:', e);
      this.isSyncingFromPreview = false;
    }
  }

  private updatePreview(domElement: HTMLElement, previewElement: HTMLElement): void {
    try {
      const html = this.options.markdownProcessor.render(this.editorContent);
      const { tocHtml, mainHtml } = this.options.markdownProcessor.extractTOC(html);
      
      // Update the preview content
      previewElement.innerHTML = mainHtml;
      
      // Update or create TOC in the editor layout (outside preview pane)
      const editorLayout = domElement.querySelector(`.${this.options.styles.editorLayout}`) as HTMLElement;
      if (editorLayout) {
        // Remove existing TOC
        const existingTOC = editorLayout.querySelector(`.${this.options.styles.tocSidebar}`);
        if (existingTOC) {
          existingTOC.remove();
        }
        
        // Add new TOC if present
        if (tocHtml) {
          editorLayout.insertAdjacentHTML('beforeend', `<aside class="${this.options.styles.tocSidebar}">${tocHtml}</aside>`);
        }
      }
      
      // Re-apply enhancements to the entire editor layout (includes TOC)
      const enhancementTarget = editorLayout || previewElement;
      
      // Add copy button functionality to code blocks
      this.options.codeBlockEnhancer.addCopyButtonFunctionality(enhancementTarget);
      
      // Update TOC position based on current property pane state (skip animation during dynamic updates)
      this.updateTOCPosition(domElement, this.options.propertyPaneDetector.isPropertyPaneOpen(), true);

      // Re-add TOC click handlers since TOC was regenerated
      this.addTOCClickHandlers(domElement);

      // Render Mermaid diagrams
      if (this.options.properties.enableMermaid) {
        console.log('🎨 EditMode: Mermaid enabled, checking for diagrams in preview...');
        console.log('🎨 EditMode: Enhancement target:', enhancementTarget);
        console.log('🎨 EditMode: Preview HTML contains mermaid:', previewElement.innerHTML.includes('mermaid'));
        
        setTimeout(async () => {
          try {
            await this.options.mermaidRenderer.renderDiagrams(enhancementTarget, this.options.styles.mermaidError);
            console.log('🎨 EditMode: Mermaid rendering completed in preview');
          } catch (e) {
            console.error('🎨 EditMode: Mermaid rendering error:', e);
          }
        }, 300);
      } else {
        console.log('🎨 EditMode: Mermaid disabled, skipping diagram rendering');
      }
    } catch (e) {
      console.error('EditMode: Preview update error:', e);
      previewElement.innerHTML = '<div class="error">Error rendering preview</div>';
    }
  }

  private updateTOCPosition(domElement: HTMLElement, isPropertyPaneOpen: boolean, skipAnimation: boolean = false): void {
    console.log('📍 EditMode: Updating TOC position, property pane open:', isPropertyPaneOpen);
    
    const tocElements = domElement.querySelectorAll('.' + this.options.styles.tocSidebar);
    console.log('📍 EditMode: Found TOC elements:', tocElements.length);
    
    tocElements.forEach((element: HTMLElement, index) => {
      // Check if state actually changed to prevent unnecessary updates
      const hasPropertyPaneClass = element.classList.contains(this.options.styles.propertyPaneOpen);
      const shouldHaveClass = isPropertyPaneOpen;
      
      if (hasPropertyPaneClass === shouldHaveClass) {
        console.log(`📍 EditMode: Element ${index} already in correct state, skipping`);
        return;
      }
      
      console.log(`📍 EditMode: Updating TOC element ${index}, current classes:`, element.className);
      
      // Temporarily disable transition if skipAnimation is true
      if (skipAnimation) {
        element.style.transition = 'none';
      }
      
      if (isPropertyPaneOpen) {
        element.classList.add(this.options.styles.propertyPaneOpen);
        console.log(`📍 EditMode: Added propertyPaneOpen class to element ${index}`);
      } else {
        element.classList.remove(this.options.styles.propertyPaneOpen);
        console.log(`📍 EditMode: Removed propertyPaneOpen class from element ${index}`);
      }
      
      // Re-enable transition after a brief delay
      if (skipAnimation) {
        setTimeout(() => {
          element.style.transition = '';
        }, 10);
      }
      
      console.log(`📍 EditMode: Element ${index} final classes:`, element.className);
    });
  }

  private renderFallbackEditor(container: HTMLElement, previewContent: HTMLElement, saveButton: HTMLButtonElement, resizer: HTMLElement, domElement: HTMLElement): void {
    console.log('📝 EditMode: Using textarea fallback editor');
    container.innerHTML = `<textarea id="markdownEditor" class="${this.options.styles.markdownEditor}" placeholder="Enter your markdown content here...">${this.editorContent}</textarea>`;
    
    const editor = container.querySelector('#markdownEditor') as HTMLTextAreaElement;
    if (!editor) return;

    // Real-time preview updates
    let updateTimeout: number;
    editor.addEventListener('input', () => {
      clearTimeout(updateTimeout);
      updateTimeout = window.setTimeout(() => {
        this.editorContent = editor.value;
        this.updatePreview(domElement, previewContent);
      }, 300);
    });

    // Save functionality
    saveButton.addEventListener('click', () => {
      this.options.onPropertyChange('markdownContent', editor.value);
      this.options.onPropertyPaneRefresh();
      saveButton.textContent = 'Saved!';
      setTimeout(() => {
        saveButton.textContent = 'Save';
      }, 2000);
    });

    // Resizable panes
    if (resizer) {
      this.initializeResizer(domElement, resizer);
    }

    // Auto-resize textarea
    this.autoResizeTextarea(editor);
  }

  private initializeResizer(domElement: HTMLElement, resizer: HTMLElement): void {
    let isResizing = false;
    let startX = 0;
    let startLeftWidth = 0;

    const handleMouseMove = (e: MouseEvent): void => {
      if (!isResizing) return;
      const dx = e.clientX - startX;
      const editorPane = domElement.querySelector(`.${this.options.styles.editorPane}`) as HTMLElement;
      const previewPane = domElement.querySelector(`.${this.options.styles.previewPane}`) as HTMLElement;

      if (editorPane && previewPane) {
        const newWidth = startLeftWidth + dx;
        const editorLayout = domElement.querySelector(`.${this.options.styles.editorLayout}`) as HTMLElement;
        const containerWidth = editorLayout ? editorLayout.offsetWidth : domElement.offsetWidth;
        const minWidth = 200;
        const maxWidth = containerWidth - 200;

        if (newWidth >= minWidth && newWidth <= maxWidth) {
          editorPane.style.width = `${newWidth}px`;
          editorPane.style.flex = 'none';
          previewPane.style.flex = 'none';
          previewPane.style.width = `${containerWidth - newWidth - 10}px`; // 10px for resizer

          // Trigger Monaco layout update
          if (this.monacoEditor) {
            this.monacoEditor.layout();
          }
        }
      }
    };

    const handleMouseUp = (): void => {
      isResizing = false;
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };

    resizer.addEventListener('mousedown', (e) => {
      isResizing = true;
      startX = e.clientX;
      const editorPane = domElement.querySelector(`.${this.options.styles.editorPane}`) as HTMLElement;
      if (editorPane) {
        startLeftWidth = editorPane.offsetWidth;
      }
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
      e.preventDefault();
    });
  }

  private autoResizeTextarea(textarea: HTMLTextAreaElement): void {
    const adjustHeight = () => {
      textarea.style.height = 'auto';
      textarea.style.height = Math.max(textarea.scrollHeight, 200) + 'px';
    };
    
    textarea.addEventListener('input', adjustHeight);
    adjustHeight(); // Initial sizing
  }

  public updateTheme(): void {
    if (this.monacoEditor) {
      MonacoLoader.setTheme(this.options.properties.theme);
    }
  }

  public dispose(): void {
    if (this.monacoEditor) {
      this.monacoEditor.dispose();
      this.monacoEditor = null;
    }
  }

  public handlePropertyPaneStateChange(isOpen: boolean): void {
    // This will be called by the main webpart when property pane state changes
    // We need to find the current dom element with the editor layout
    const editorLayout = document.querySelector(`.${this.options.styles.editorLayout}`) as HTMLElement;
    if (editorLayout) {
      this.updateTOCPosition(editorLayout.parentElement || document.body, isOpen, false); // Allow animation for manual property pane changes
    }
  }

  private toggleEditorCollapse(domElement: HTMLElement, toggleButton: HTMLButtonElement): void {
    const editorPane = domElement.querySelector(`.${this.options.styles.editorPane}`) as HTMLElement;
    const previewPane = domElement.querySelector(`.${this.options.styles.previewPane}`) as HTMLElement;
    const resizer = domElement.querySelector(`.${this.options.styles.resizer}`) as HTMLElement;
    const monacoContainer = domElement.querySelector('#monacoContainer') as HTMLElement;
    const editorHeader = domElement.querySelector(`.${this.options.styles.editorHeader}`) as HTMLElement;
    const headerTitle = editorHeader?.querySelector('h3') as HTMLElement;
    const importButton = domElement.querySelector('#importMdFile') as HTMLButtonElement;
    const saveButton = domElement.querySelector('#saveContent') as HTMLButtonElement;

    if (!editorPane || !previewPane) return;

    this.isEditorCollapsed = !this.isEditorCollapsed;

    if (this.isEditorCollapsed) {
      // Collapse editor
      editorPane.style.flex = 'none';
      editorPane.style.width = '50px';
      editorPane.style.minWidth = '50px';
      previewPane.style.flex = '1';
      previewPane.style.width = 'auto';
      if (resizer) resizer.style.display = 'none';
      if (monacoContainer) monacoContainer.style.display = 'none';

      // Make header vertical with icon-only buttons
      if (editorHeader) {
        editorHeader.style.flexDirection = 'column';
        editorHeader.style.alignItems = 'center';
        editorHeader.style.padding = '12px 4px';
        editorHeader.style.gap = '8px';
      }
      if (headerTitle) {
        headerTitle.style.writingMode = 'vertical-rl';
        headerTitle.style.textOrientation = 'mixed';
        headerTitle.style.fontSize = '12px';
        headerTitle.style.whiteSpace = 'nowrap';
      }

      // Update buttons to icon-only
      if (importButton) importButton.textContent = '📁';
      if (toggleButton) {
        toggleButton.textContent = '▶️';
        toggleButton.title = 'Expand editor';
      }
      if (saveButton) saveButton.textContent = '💾';
    } else {
      // Expand editor
      editorPane.style.flex = '1';
      editorPane.style.width = '';
      editorPane.style.minWidth = '200px';
      previewPane.style.flex = '1';
      previewPane.style.width = '';
      if (resizer) resizer.style.display = 'block';
      if (monacoContainer) monacoContainer.style.display = 'flex';

      // Restore header to horizontal
      if (editorHeader) {
        editorHeader.style.flexDirection = '';
        editorHeader.style.alignItems = '';
        editorHeader.style.padding = '';
        editorHeader.style.gap = '';
      }
      if (headerTitle) {
        headerTitle.style.writingMode = '';
        headerTitle.style.textOrientation = '';
        headerTitle.style.fontSize = '';
        headerTitle.style.whiteSpace = '';
      }

      // Restore button text
      if (importButton) importButton.textContent = '📁 Import';
      if (toggleButton) {
        toggleButton.textContent = '◀️ Collapse';
        toggleButton.title = 'Collapse editor';
      }
      if (saveButton) saveButton.textContent = '💾 Save';
    }

    // Trigger Monaco layout update
    if (this.monacoEditor) {
      setTimeout(() => {
        this.monacoEditor.layout();
      }, 300);
    }
  }

  private importMarkdownFile(): void {
    // Create a hidden file input
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.md,.markdown,.txt';
    fileInput.style.display = 'none';

    fileInput.addEventListener('change', (event: Event) => {
      const target = event.target as HTMLInputElement;
      const file = target.files?.[0];

      if (!file) return;

      const reader = new FileReader();
      reader.onload = (e: ProgressEvent<FileReader>) => {
        const content = e.target?.result as string;
        if (content && this.monacoEditor) {
          this.monacoEditor.setValue(content);
          this.editorContent = content;
        }
      };
      reader.readAsText(file);
    });

    // Trigger file selection dialog
    document.body.appendChild(fileInput);
    fileInput.click();
    document.body.removeChild(fileInput);
  }

  private async exportToPdf(tocHtml: string, mainHtml: string): Promise<void> {
    try {
      // Get the current rendered content from the preview
      const previewContent = document.querySelector('#previewContent') as HTMLElement;
      if (previewContent) {
        mainHtml = previewContent.innerHTML;
      }

      // Get TOC if present
      const tocSidebar = document.querySelector(`.${this.options.styles.tocSidebar}`) as HTMLElement;
      if (tocSidebar) {
        tocHtml = tocSidebar.innerHTML;
      }

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