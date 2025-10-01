import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

import styles from './BetterMarkdownWebPart.module.scss';
import * as strings from 'BetterMarkdownWebPartStrings';

// Import highlight.js CSS for syntax highlighting
import 'highlight.js/styles/atom-one-dark.css';

// Import utilities
import { MermaidRenderer } from './utils/MermaidRenderer';
import { MarkdownProcessor, IMarkdownProcessorOptions } from './utils/MarkdownProcessor';
import { PropertyPaneDetector } from './utils/PropertyPaneDetector';
import { CodeBlockEnhancer } from './utils/CodeBlockEnhancer';
import { KaTeXLoader } from './utils/KaTeXLoader';
import { EditModeManager } from './utils/EditModeManager';

export interface IBetterMarkdownWebPartProps {
  markdownContent: string;
  enableMermaid: boolean;
  enableMath: boolean;
  enableTOC: boolean;
  enableSyntaxHighlighting: boolean;
  theme: string;
}

export default class BetterMarkdownWebPart extends BaseClientSideWebPart<IBetterMarkdownWebPartProps> {
  private mermaidRenderer: MermaidRenderer;
  private markdownProcessor: MarkdownProcessor;
  private propertyPaneDetector: PropertyPaneDetector;
  private codeBlockEnhancer: CodeBlockEnhancer;
  private editModeManager: EditModeManager;
  private isEditorMode: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      // Set default properties if not set
      this.ensureDefaultProperties();

      // Load KaTeX CSS if math is enabled
      if (this.properties.enableMath && !KaTeXLoader.isKatexCssLoaded) {
        KaTeXLoader.loadKatexCss();
      }

      // Initialize utilities
      this.initializeUtilities();
      
      // Initialize mermaid if enabled
      if (this.properties.enableMermaid) {
        console.log('🚀 WebPart: Initializing Mermaid...');
        this.mermaidRenderer.initialize(this.properties.theme).catch(e => {
          console.error('🚀 WebPart: Mermaid initialization failed:', e);
        });
      } else {
        console.log('🚀 WebPart: Mermaid disabled in properties');
      }
    });
  }

  private ensureDefaultProperties(): void {
    // Set default values for properties if they're undefined
    if (this.properties.enableMermaid === undefined) {
      this.properties.enableMermaid = true;
    }
    if (this.properties.enableMath === undefined) {
      this.properties.enableMath = true;
    }
    if (this.properties.enableTOC === undefined) {
      this.properties.enableTOC = true;
    }
    if (this.properties.enableSyntaxHighlighting === undefined) {
      this.properties.enableSyntaxHighlighting = true;
    }
    if (!this.properties.theme) {
      this.properties.theme = 'light';
    }
    if (!this.properties.markdownContent) {
      this.properties.markdownContent = '# Better Markdown\n\nStart editing to see the preview...';
    }

    console.log('🔧 WebPart: Properties initialized:', {
      enableMermaid: this.properties.enableMermaid,
      enableTOC: this.properties.enableTOC,
      enableMath: this.properties.enableMath,
      enableSyntaxHighlighting: this.properties.enableSyntaxHighlighting,
      theme: this.properties.theme
    });
  }

  private initializeUtilities(): void {
    // Initialize Mermaid renderer
    this.mermaidRenderer = new MermaidRenderer();

    // Initialize markdown processor with options
    const processorOptions: IMarkdownProcessorOptions = {
      enableSyntaxHighlighting: this.properties.enableSyntaxHighlighting,
      enableTOC: this.properties.enableTOC,
      enableMath: this.properties.enableMath,
      enableMermaid: this.properties.enableMermaid,
      theme: this.properties.theme,
      styleClasses: {
        codeToolbar: styles.codeToolbar,
        toolbar: styles.toolbar,
        toolbarItem: styles.toolbarItem,
        codeLine: styles.codeLine,
        lineNumber: styles.lineNumber,
        lineContent: styles.lineContent,
        tableOfContents: styles.tableOfContents,
        blockquote: styles.blockquote,
        mathBlock: styles.mathBlock,
        mathError: styles.mathError,
        mermaidContainer: styles.mermaidContainer,
        mermaidError: styles.mermaidError
      }
    };
    this.markdownProcessor = new MarkdownProcessor(processorOptions);

    // Initialize code block enhancer
    this.codeBlockEnhancer = new CodeBlockEnhancer({
      codeToolbar: styles.codeToolbar,
      lineContent: styles.lineContent
    });

    // Initialize property pane detector
    this.propertyPaneDetector = new PropertyPaneDetector((isOpen: boolean) => {
      this.updateTOCPosition(isOpen);
      
      // Also notify edit mode manager if we're in edit mode
      if (this.isEditorMode && this.editModeManager) {
        void this.editModeManager.handlePropertyPaneStateChange(isOpen);
      }
    });
    this.propertyPaneDetector.start();

    // Initialize edit mode manager
    this.editModeManager = new EditModeManager({
      styles: styles,
      properties: this.properties,
      markdownProcessor: this.markdownProcessor,
      mermaidRenderer: this.mermaidRenderer,
      codeBlockEnhancer: this.codeBlockEnhancer,
      propertyPaneDetector: this.propertyPaneDetector,
      onPropertyChange: (property: string, value: any) => {
        (this.properties as any)[property] = value;
        this.onPropertyPaneFieldChanged(property, '', value);
      },
      onPropertyPaneRefresh: () => {
        this.context.propertyPane.refresh();
      }
    });
  }

  private updateTOCPosition(isPropertyPaneOpen: boolean, skipAnimation: boolean = false): void {
    console.log('📍 WebPart: Updating TOC position, property pane open:', isPropertyPaneOpen);
    
    const tocElements = this.domElement.querySelectorAll('.' + styles.tocSidebar);
    console.log('📍 WebPart: Found TOC elements:', tocElements.length);
    
    tocElements.forEach((element: HTMLElement, index) => {
      // Check if state actually changed to prevent unnecessary updates
      const hasPropertyPaneClass = element.classList.contains(styles.propertyPaneOpen);
      const shouldHaveClass = isPropertyPaneOpen;
      
      if (hasPropertyPaneClass === shouldHaveClass) {
        console.log(`📍 WebPart: Element ${index} already in correct state, skipping`);
        return;
      }
      
      console.log(`📍 WebPart: Updating TOC element ${index}, current classes:`, element.className);
      
      // Temporarily disable transition if skipAnimation is true
      if (skipAnimation) {
        element.style.transition = 'none';
      }
      
      if (isPropertyPaneOpen) {
        element.classList.add(styles.propertyPaneOpen);
        console.log(`📍 WebPart: Added propertyPaneOpen class to element ${index}`);
      } else {
        element.classList.remove(styles.propertyPaneOpen);
        console.log(`📍 WebPart: Removed propertyPaneOpen class from element ${index}`);
      }
      
      // Re-enable transition after a brief delay
      if (skipAnimation) {
        setTimeout(() => {
          element.style.transition = '';
        }, 10);
      }
      
      console.log(`📍 WebPart: Element ${index} final classes:`, element.className);
    });
  }

  public render(): void {
    // Detect if we're in edit mode
    this.isEditorMode = this.displayMode === DisplayMode.Edit;
    
    // Update theme for renderers
    this.mermaidRenderer.updateTheme(this.properties.theme);
    if (this.editModeManager) {
      this.editModeManager.updateTheme();
    }
    
    if (this.isEditorMode) {
      this.editModeManager.renderEditMode(this.domElement);
    } else {
      this.renderViewMode();
    }
  }

  private renderViewMode(): void {
    const content = this.properties.markdownContent || '# Wiki.js Style Markdown\n\nStart editing to see the preview...';

    // Render markdown content
    const html = this.markdownProcessor.render(content);

    // Extract TOC if present and TOC is enabled
    const { tocHtml, mainHtml } = this.markdownProcessor.extractTOC(html);

    // Create layout with sticky TOC and export button
    this.domElement.innerHTML = `<div class="${styles.betterMarkdown} ${styles[this.properties.theme] || ''}">
      <div class="${styles.exportActions}">
        <button id="exportPdf" class="${styles.exportButton}" title="Export as PDF">📄 Export PDF</button>
      </div>
      <div class="${styles.layout}">
        <div class="${styles.mainContent}">
          ${mainHtml}
        </div>
        ${tocHtml ? `<aside class="${styles.tocSidebar}">${tocHtml}</aside>` : ''}
      </div>
    </div>`;

    // Add export button functionality
    const exportButton = this.domElement.querySelector('#exportPdf') as HTMLButtonElement;
    if (exportButton) {
      exportButton.addEventListener('click', () => {
        void this.exportToPdf(tocHtml, mainHtml);
      });
    }

    // Post-render enhancements
    this.enhanceRenderedContent();
  }



  private async enhanceRenderedContent(): Promise<void> {
    console.log('🎨 WebPart: Starting content enhancement...');
    
    // Add copy button functionality to code blocks
    this.codeBlockEnhancer.addCopyButtonFunctionality(this.domElement);

    // Update TOC position based on current property pane state
    this.updateTOCPosition(this.propertyPaneDetector.isPropertyPaneOpen());

    // Render Mermaid diagrams - do this last and with a delay to ensure everything is ready
    if (this.properties.enableMermaid) {
      console.log('🎨 WebPart: Mermaid enabled, attempting to render diagrams...');
      
      // Add a small delay to ensure DOM is fully updated and mermaid is loaded
      setTimeout(() => {
        void (async () => {
          try {
            await this.mermaidRenderer.renderDiagrams(this.domElement, styles.mermaidError);
            console.log('🎨 WebPart: Mermaid rendering completed');
          } catch (e) {
            console.error('🎨 WebPart: Mermaid rendering error:', e);
          }
        })();
      }, 500); // 500ms delay to ensure mermaid is loaded
    } else {
      console.log('🎨 WebPart: Mermaid disabled, skipping diagram rendering');
    }
    
    console.log('🎨 WebPart: Content enhancement completed');
  }

  private async exportToPdf(tocHtml: string, mainHtml: string): Promise<void> {
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

  protected onDispose(): void {
    // Clean up utilities
    if (this.propertyPaneDetector) {
      this.propertyPaneDetector.stop();
    }
    if (this.codeBlockEnhancer) {
      this.codeBlockEnhancer.removeEnhancedMarkers(this.domElement);
    }
    if (this.editModeManager) {
      this.editModeManager.dispose();
    }
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log('🔧 WebPart: Property changed:', propertyPath, 'from', oldValue, 'to', newValue);
    
    // Handle Mermaid enable/disable
    if (propertyPath === 'enableMermaid') {
      if (newValue && !this.mermaidRenderer.initialized) {
        console.log('🚀 WebPart: Initializing Mermaid due to property change...');
        this.mermaidRenderer.initialize(this.properties.theme).catch(e => {
          console.error('🚀 WebPart: Mermaid initialization failed:', e);
        });
      }
    }
    
    // Handle theme changes
    if (propertyPath === 'theme') {
      this.mermaidRenderer.updateTheme(newValue);
      if (this.editModeManager) {
        this.editModeManager.updateTheme();
      }
    }
    
    // Update markdown processor options
    if (['enableSyntaxHighlighting', 'enableTOC', 'enableMath', 'enableMermaid', 'theme'].includes(propertyPath)) {
      this.markdownProcessor.updateOptions({
        enableSyntaxHighlighting: this.properties.enableSyntaxHighlighting,
        enableTOC: this.properties.enableTOC,
        enableMath: this.properties.enableMath,
        enableMermaid: this.properties.enableMermaid,
        theme: this.properties.theme
      });
    }
    
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('markdownContent', {
                  label: 'Markdown Content',
                  multiline: true,
                  rows: 25,
                  description: 'Enter your markdown content here'
                }),
                PropertyPaneToggle('enableMermaid', {
                  label: 'Enable Mermaid Diagrams',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('enableMath', {
                  label: 'Enable Math (KaTeX)',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('enableTOC', {
                  label: 'Enable Table of Contents',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('enableSyntaxHighlighting', {
                  label: 'Enable Syntax Highlighting',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('theme', {
                  label: 'Theme',
                  description: 'Enter: light or dark'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}