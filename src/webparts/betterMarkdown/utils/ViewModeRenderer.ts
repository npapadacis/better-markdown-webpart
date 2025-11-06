import { MarkdownProcessor } from './MarkdownProcessor';
import { MermaidRenderer } from './MermaidRenderer';
import { CodeBlockEnhancer } from './CodeBlockEnhancer';
import { PropertyPaneDetector } from './PropertyPaneDetector';
import { IFileMetadata } from './SharePointService';

export interface IViewModeRendererOptions {
  styles: any;
  markdownProcessor: MarkdownProcessor;
  mermaidRenderer: MermaidRenderer;
  codeBlockEnhancer: CodeBlockEnhancer;
  propertyPaneDetector: PropertyPaneDetector;
  enableMermaid: boolean;
  enableVersionHistory: boolean;
  theme: string;
  onExportPdf: (tocHtml: string, mainHtml: string) => Promise<void>;
  onReloadFile: () => Promise<void>;
  onShowVersionHistory: () => Promise<void>;
}

export class ViewModeRenderer {
  private options: IViewModeRendererOptions;

  constructor(options: IViewModeRendererOptions) {
    this.options = options;
  }

  /**
   * Render view mode with markdown content
   */
  public renderViewMode(
    domElement: HTMLElement,
    content: string,
    fileMetadata?: IFileMetadata
  ): void {
    const displayContent = content || '# Wiki.js Style Markdown\n\nStart editing to see the preview...';

    // Render markdown content
    const html = this.options.markdownProcessor.render(displayContent);

    // Extract TOC if present and TOC is enabled
    const { tocHtml, mainHtml } = this.options.markdownProcessor.extractTOC(html);

    // Build enhanced file info display if file is selected
    const fileInfo = this.buildFileInfoHtml(fileMetadata);

    // Create layout with sticky TOC and export button
    domElement.innerHTML = `<div class="${this.options.styles.betterMarkdown} ${this.options.styles[this.options.theme] || ''}">
      <div class="${this.options.styles.exportActions}">
        ${fileInfo}
        <button id="exportPdf" class="${this.options.styles.exportButton}" title="Export as PDF">📄 Export PDF</button>
      </div>
      <div class="${this.options.styles.layout}">
        <div class="${this.options.styles.mainContent}">
          ${mainHtml}
        </div>
        ${tocHtml ? `<aside class="${this.options.styles.tocSidebar}">${tocHtml}</aside>` : ''}
      </div>
    </div>`;

    // Set up event handlers
    this.setupEventHandlers(domElement, tocHtml, mainHtml);

    // Post-render enhancements
    void this.enhanceRenderedContent(domElement);
  }

  private buildFileInfoHtml(fileMetadata?: IFileMetadata): string {
    if (!fileMetadata) return '';

    return `<div class="${this.options.styles.fileInfo}">
      <div class="${this.options.styles.fileInfoMain}">
        <span class="${this.options.styles.fileInfoLabel}">📄 Source:</span>
        <span class="${this.options.styles.fileName}">${fileMetadata.name || 'Unknown'}</span>
      </div>
      <div class="${this.options.styles.fileMetadata}">
        <span class="${this.options.styles.metadataItem}">
          <span class="${this.options.styles.metadataLabel}">Modified:</span>
          <span class="${this.options.styles.metadataValue}">${new Date(fileMetadata.timeLastModified).toLocaleString()}</span>
        </span>
        <span class="${this.options.styles.metadataItem}">
          <span class="${this.options.styles.metadataLabel}">By:</span>
          <span class="${this.options.styles.metadataValue}">${fileMetadata.author}</span>
        </span>
        <span class="${this.options.styles.metadataItem}">
          <span class="${this.options.styles.metadataLabel}">Size:</span>
          <span class="${this.options.styles.metadataValue}">${this.formatFileSize(fileMetadata.length)}</span>
        </span>
        ${this.options.enableVersionHistory ? `
          <button id="viewHistory" class="${this.options.styles.historyButton}" title="View version history">📜 History</button>
        ` : ''}
      </div>
      <button id="reloadFile" class="${this.options.styles.reloadButton}" title="Reload from SharePoint">🔄 Reload</button>
    </div>`;
  }

  private setupEventHandlers(
    domElement: HTMLElement,
    tocHtml: string,
    mainHtml: string
  ): void {
    // Add export button functionality
    const exportButton = domElement.querySelector('#exportPdf') as HTMLButtonElement;
    if (exportButton) {
      exportButton.addEventListener('click', () => {
        void this.options.onExportPdf(tocHtml, mainHtml);
      });
    }

    // Add reload button functionality
    const reloadButton = domElement.querySelector('#reloadFile') as HTMLButtonElement;
    if (reloadButton) {
      reloadButton.addEventListener('click', () => {
        void this.options.onReloadFile();
      });
    }

    // Add version history button functionality
    const historyButton = domElement.querySelector('#viewHistory') as HTMLButtonElement;
    if (historyButton) {
      historyButton.addEventListener('click', () => {
        void this.options.onShowVersionHistory();
      });
    }
  }

  private async enhanceRenderedContent(domElement: HTMLElement): Promise<void> {
    console.log('🎨 ViewMode: Starting content enhancement...');

    // Add copy button functionality to code blocks
    this.options.codeBlockEnhancer.addCopyButtonFunctionality(domElement);

    // Update TOC position based on current property pane state
    this.updateTOCPosition(domElement, this.options.propertyPaneDetector.isPropertyPaneOpen());

    // Render Mermaid diagrams - do this last and with a delay to ensure everything is ready
    if (this.options.enableMermaid) {
      console.log('🎨 ViewMode: Mermaid enabled, attempting to render diagrams...');

      // Add a small delay to ensure DOM is fully updated and mermaid is loaded
      setTimeout(() => {
        void (async () => {
          try {
            // Ensure mermaid is initialized before rendering
            if (!this.options.mermaidRenderer.initialized) {
              console.log('🎨 ViewMode: Mermaid not initialized, initializing now...');
              await this.options.mermaidRenderer.initialize(this.options.theme);
            }
            await this.options.mermaidRenderer.renderDiagrams(domElement, this.options.styles.mermaidError);
            console.log('🎨 ViewMode: Mermaid rendering completed');
          } catch (e) {
            console.error('🎨 ViewMode: Mermaid rendering error:', e);
          }
        })();
      }, 500); // 500ms delay to ensure mermaid is loaded
    } else {
      console.log('🎨 ViewMode: Mermaid disabled, skipping diagram rendering');
    }

    console.log('🎨 ViewMode: Content enhancement completed');
  }

  public updateTOCPosition(domElement: HTMLElement, isPropertyPaneOpen: boolean, skipAnimation: boolean = false): void {
    console.log('📍 ViewMode: Updating TOC position, property pane open:', isPropertyPaneOpen);

    const tocElements = domElement.querySelectorAll('.' + this.options.styles.tocSidebar);
    console.log('📍 ViewMode: Found TOC elements:', tocElements.length);

    tocElements.forEach((element: HTMLElement, index) => {
      // Check if state actually changed to prevent unnecessary updates
      const hasPropertyPaneClass = element.classList.contains(this.options.styles.propertyPaneOpen);
      const shouldHaveClass = isPropertyPaneOpen;

      if (hasPropertyPaneClass === shouldHaveClass) {
        console.log(`📍 ViewMode: Element ${index} already in correct state, skipping`);
        return;
      }

      console.log(`📍 ViewMode: Updating TOC element ${index}, current classes:`, element.className);

      // Temporarily disable transition if skipAnimation is true
      if (skipAnimation) {
        element.style.transition = 'none';
      }

      if (isPropertyPaneOpen) {
        element.classList.add(this.options.styles.propertyPaneOpen);
        console.log(`📍 ViewMode: Added propertyPaneOpen class to element ${index}`);
      } else {
        element.classList.remove(this.options.styles.propertyPaneOpen);
        console.log(`📍 ViewMode: Removed propertyPaneOpen class from element ${index}`);
      }

      // Re-enable transition after a brief delay
      if (skipAnimation) {
        setTimeout(() => {
          element.style.transition = '';
        }, 10);
      }

      console.log(`📍 ViewMode: Element ${index} final classes:`, element.className);
    });
  }

  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
  }
}
