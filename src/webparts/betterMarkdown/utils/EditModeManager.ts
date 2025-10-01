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
              <button id="saveContent" class="${this.options.styles.saveButton}">Save</button>
            </div>
          </div>
          <div id="monacoContainer" class="${this.options.styles.monacoContainer}"></div>
        </div>
        <div class="${this.options.styles.resizer}"></div>
        <div class="${this.options.styles.previewPane}">
          <div class="${this.options.styles.previewHeader}">
            <h3>Live Preview</h3>
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

  private async initializeMonacoEditor(domElement: HTMLElement): Promise<void> {
    const container = domElement.querySelector('#monacoContainer') as HTMLElement;
    const previewContent = domElement.querySelector('#previewContent') as HTMLElement;
    const saveButton = domElement.querySelector('#saveContent') as HTMLButtonElement;
    const resizer = domElement.querySelector(`.${this.options.styles.resizer}`) as HTMLElement;
    
    if (!container || !previewContent || !saveButton) return;

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

      // Save functionality
      saveButton.addEventListener('click', () => {
        const content = this.monacoEditor.getValue();
        this.options.onPropertyChange('markdownContent', content);
        this.options.onPropertyPaneRefresh();
        saveButton.textContent = 'Saved!';
        setTimeout(() => {
          saveButton.textContent = 'Save';
        }, 2000);
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

    resizer.addEventListener('mousedown', (e) => {
      isResizing = true;
      startX = e.clientX;
      const editorPane = domElement.querySelector(`.${this.options.styles.editorPane}`) as HTMLElement;
      if (editorPane) {
        startLeftWidth = parseInt(window.getComputedStyle(editorPane).width, 10);
      }
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
      e.preventDefault();
    });

    const handleMouseMove = (e: MouseEvent) => {
      if (!isResizing) return;
      const dx = e.clientX - startX;
      const editorPane = domElement.querySelector(`.${this.options.styles.editorPane}`) as HTMLElement;
      const previewPane = domElement.querySelector(`.${this.options.styles.previewPane}`) as HTMLElement;
      
      if (editorPane && previewPane) {
        const newWidth = startLeftWidth + dx;
        const containerWidth = domElement.offsetWidth;
        const minWidth = 200;
        const maxWidth = containerWidth - 200;
        
        if (newWidth >= minWidth && newWidth <= maxWidth) {
          editorPane.style.width = `${newWidth}px`;
          previewPane.style.width = `${containerWidth - newWidth - 10}px`; // 10px for resizer
        }
      }
    };

    const handleMouseUp = () => {
      isResizing = false;
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
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
}