import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
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
import { SharePointService, IFileMetadata } from './utils/SharePointService';
import { PdfExportManager } from './utils/PdfExportManager';
import { FileVersionManager } from './utils/FileVersionManager';
import { FileSourceManager } from './utils/FileSourceManager';
import { ViewModeRenderer } from './utils/ViewModeRenderer';

export interface IBetterMarkdownWebPartProps {
  markdownContent: string;
  contentSource: 'browser' | 'url' | 'manual'; // New: source selection
  fileUrl: string; // New: direct URL input
  selectedLibrary: string;
  selectedFolder: string;
  selectedMarkdownFile: string;
  enableAutoRefresh: boolean;
  enableVersionHistory: boolean;
  enableMermaid: boolean;
  enableMath: boolean;
  enableTOC: boolean;
  enableSyntaxHighlighting: boolean;
  theme: string;
  lastModified?: string;
  fileMetadata?: IFileMetadata;
}

export default class BetterMarkdownWebPart extends BaseClientSideWebPart<IBetterMarkdownWebPartProps> {
  private mermaidRenderer: MermaidRenderer;
  private markdownProcessor: MarkdownProcessor;
  private propertyPaneDetector: PropertyPaneDetector;
  private codeBlockEnhancer: CodeBlockEnhancer;
  private editModeManager: EditModeManager;
  private sharePointService: SharePointService;
  private fileVersionManager: FileVersionManager;
  private fileSourceManager: FileSourceManager;
  private viewModeRenderer: ViewModeRenderer;
  private isEditorMode: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
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

      // Load SharePoint library options for property pane
      void this.fileSourceManager.loadLibraryOptions(this.properties.selectedLibrary);

      // Load folders and files if library is already selected
      if (this.properties.selectedLibrary) {
        void this.fileSourceManager.loadFolderOptions(this.properties.selectedLibrary).then(() => {
          if (this.properties.selectedFolder !== undefined) {
            void this.fileSourceManager.loadFileOptions(this.properties.selectedLibrary, this.properties.selectedFolder);
          }
        });
      }

      // Always load latest file content on page load if a source is configured
      if (this.properties.contentSource === 'browser' && this.properties.selectedMarkdownFile) {
        console.log('🔄 WebPart Init: Loading latest version of selected file from SharePoint...');
        await this.loadSelectedFile();
      } else if (this.properties.contentSource === 'url' && this.properties.fileUrl) {
        console.log('🔄 WebPart Init: Loading from URL...');
        await this.loadFromUrl();
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
    if (this.properties.enableAutoRefresh === undefined) {
      this.properties.enableAutoRefresh = false;
    }
    if (this.properties.enableVersionHistory === undefined) {
      this.properties.enableVersionHistory = true;
    }
    if (!this.properties.theme) {
      this.properties.theme = 'light';
    }
    if (!this.properties.contentSource) {
      this.properties.contentSource = 'manual';
    }
    if (!this.properties.markdownContent) {
      this.properties.markdownContent = '# Better Markdown\n\nStart editing to see the preview...';
    }

    console.log('🔧 WebPart: Properties initialized:', {
      enableMermaid: this.properties.enableMermaid,
      enableTOC: this.properties.enableTOC,
      enableMath: this.properties.enableMath,
      enableSyntaxHighlighting: this.properties.enableSyntaxHighlighting,
      enableAutoRefresh: this.properties.enableAutoRefresh,
      enableVersionHistory: this.properties.enableVersionHistory,
      theme: this.properties.theme,
      contentSource: this.properties.contentSource
    });
  }

  private initializeUtilities(): void {
    // Initialize SharePoint service
    this.sharePointService = new SharePointService(this.context);

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
      },
      onSaveToSharePoint: async (content: string) => {
        return await this.saveToSharePoint(content);
      }
    });

    // Initialize file version manager
    this.fileVersionManager = new FileVersionManager({
      styles: styles,
      sharePointService: this.sharePointService,
      onSave: async (content: string) => await this.saveToSharePoint(content),
      onVersionRestored: () => this.render()
    });

    // Initialize file source manager
    this.fileSourceManager = new FileSourceManager({
      sharePointService: this.sharePointService,
      onContentLoaded: (content: string, metadata: IFileMetadata | null, lastModified?: string) => {
        this.properties.markdownContent = content;
        this.properties.fileMetadata = metadata;
        this.properties.lastModified = lastModified;
        this.render();
      },
      onAutoRefreshSetup: (fileUrl: string, callback: () => void) => {
        this.sharePointService.subscribeToFileChanges(fileUrl, callback);
      }
    });

    // Initialize view mode renderer
    this.viewModeRenderer = new ViewModeRenderer({
      styles: styles,
      markdownProcessor: this.markdownProcessor,
      mermaidRenderer: this.mermaidRenderer,
      codeBlockEnhancer: this.codeBlockEnhancer,
      propertyPaneDetector: this.propertyPaneDetector,
      enableMermaid: this.properties.enableMermaid,
      enableVersionHistory: this.properties.enableVersionHistory,
      theme: this.properties.theme,
      onExportPdf: async (tocHtml, mainHtml) => await PdfExportManager.exportToPdf(tocHtml, mainHtml),
      onReloadFile: async () => await this.loadSelectedFile(),
      onShowVersionHistory: async () => await this.showVersionHistory()
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

    // Update view mode renderer options
    if (this.viewModeRenderer) {
      this.viewModeRenderer = new ViewModeRenderer({
        styles: styles,
        markdownProcessor: this.markdownProcessor,
        mermaidRenderer: this.mermaidRenderer,
        codeBlockEnhancer: this.codeBlockEnhancer,
        propertyPaneDetector: this.propertyPaneDetector,
        enableMermaid: this.properties.enableMermaid,
        enableVersionHistory: this.properties.enableVersionHistory,
        theme: this.properties.theme,
        onExportPdf: async (tocHtml, mainHtml) => await PdfExportManager.exportToPdf(tocHtml, mainHtml),
        onReloadFile: async () => await this.loadSelectedFile(),
        onShowVersionHistory: async () => await this.showVersionHistory()
      });
    }

    if (this.isEditorMode) {
      this.editModeManager.renderEditMode(this.domElement);
    } else {
      this.viewModeRenderer.renderViewMode(
        this.domElement,
        this.properties.markdownContent,
        this.properties.fileMetadata
      );
    }
  }



  private getServerRelativeUrl(absoluteUrl: string): string {
    try {
      const url = new URL(absoluteUrl);
      return url.pathname;
    } catch {
      return absoluteUrl;
    }
  }

  private async checkAndReloadFile(fileUrl: string): Promise<void> {
    try {
      console.log('📄 Auto-refresh triggered, checking for unsaved changes...');

      // Check if we're in edit mode and have unsaved changes
      if (this.isEditorMode && this.editModeManager && this.editModeManager.hasUnsavedEdits()) {
        console.log('⚠️ Auto-refresh blocked: User has unsaved changes in editor');

        // Fetch the new content but don't auto-apply it
        const content = await this.sharePointService.getFileContent(fileUrl);
        const metadata = await this.sharePointService.getFileMetadata(fileUrl);

        // Let the editor handle the conflict
        this.editModeManager.handleExternalFileChange(content);

        // Update metadata if user accepted the change
        if (!this.editModeManager.hasUnsavedEdits()) {
          this.properties.markdownContent = content;
          this.properties.fileMetadata = metadata;
          this.properties.lastModified = metadata?.timeLastModified;
        }
        return;
      }

      // Safe to reload (view mode or no unsaved changes)
      console.log('📄 Reloading file...');
      const content = await this.sharePointService.getFileContent(fileUrl);
      const metadata = await this.sharePointService.getFileMetadata(fileUrl);

      this.properties.markdownContent = content;
      this.properties.fileMetadata = metadata;
      this.properties.lastModified = metadata?.timeLastModified;

      console.log('📄 File reloaded successfully, re-rendering...');
      this.render();
    } catch (error) {
      console.error('Error reloading file:', error);
    }
  }

  private async loadSelectedFile(): Promise<void> {
    await this.fileSourceManager.loadSelectedFile(
      this.properties.selectedMarkdownFile,
      this.properties.enableAutoRefresh,
      () => this.checkAndReloadFile(this.properties.selectedMarkdownFile)
    );
  }

  private async loadFromUrl(): Promise<void> {
    await this.fileSourceManager.loadFromUrl(this.properties.fileUrl);
  }

  private async saveToSharePoint(content: string): Promise<boolean> {
    try {
      // Determine the file URL to save to
      let fileUrl: string | undefined;

      if (this.properties.selectedMarkdownFile) {
        // From library browser dropdown
        fileUrl = this.properties.selectedMarkdownFile;
      }

      if (!fileUrl) {
        console.warn('No SharePoint file selected to save to');
        return false;
      }

      // Check if file was modified since last load (conflict detection)
      if (this.properties.lastModified) {
        const isModified = await this.sharePointService.checkFileModified(fileUrl, this.properties.lastModified);
        if (isModified) {
          const overwrite = confirm(
            'Warning: This file has been modified by someone else since you opened it.\n\n' +
            'Click OK to overwrite their changes, or Cancel to avoid losing their work.'
          );
          if (!overwrite) {
            return false;
          }
        }
      }

      // Save to SharePoint
      const success = await this.sharePointService.saveFileContent(fileUrl, content);

      if (success) {
        // Update last modified timestamp
        const metadata = await this.sharePointService.getFileMetadata(fileUrl);
        this.properties.lastModified = metadata?.timeLastModified;
        this.properties.fileMetadata = metadata;
      }

      return success;
    } catch (error) {
      console.error('Error saving to SharePoint:', error);
      alert(`Failed to save to SharePoint: ${error.message}`);
      return false;
    }
  }

  private async showVersionHistory(): Promise<void> {
    await this.fileVersionManager.showVersionHistory(this.properties.selectedMarkdownFile);
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
    if (this.sharePointService) {
      this.sharePointService.unsubscribeFromFileChanges();
    }
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log('🔧 WebPart: Property changed:', propertyPath, 'from', oldValue, 'to', newValue);

    // Check for unsaved changes in edit mode before making changes
    if (this.isEditorMode && this.editModeManager && this.editModeManager.hasUnsavedEdits()) {
      const message = '⚠️ WARNING: You have unsaved changes in the editor.\n\nChanging property pane settings will discard your unsaved changes.\n\nClick OK to discard changes and continue, or Cancel to keep editing.';
      if (!confirm(message)) {
        // User cancelled, revert the property change
        this.properties[propertyPath] = oldValue;
        this.context.propertyPane.refresh();
        return;
      }
    }

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
        this.editModeManager.resetEditorState();
      }
    }

    // Handle SharePoint cascading dropdowns
    if (propertyPath === 'selectedLibrary') {
      this.properties.selectedFolder = '';
      this.properties.selectedMarkdownFile = '';
      void this.fileSourceManager.loadFolderOptions(this.properties.selectedLibrary).then(() => {
        void this.fileSourceManager.loadFileOptions(this.properties.selectedLibrary, this.properties.selectedFolder).then(() => {
          this.context.propertyPane.refresh();
          this.render();
        });
      });
    }

    if (propertyPath === 'selectedFolder') {
      this.properties.selectedMarkdownFile = '';
      void this.fileSourceManager.loadFileOptions(this.properties.selectedLibrary, this.properties.selectedFolder).then(() => {
        this.context.propertyPane.refresh();
        this.render();
      });
    }

    if (propertyPath === 'selectedMarkdownFile') {
      void this.loadSelectedFile();
    }

    if (propertyPath === 'contentSource') {
      // Clear fileUrl if switching away from URL mode
      if (newValue !== 'url') {
        this.properties.fileUrl = '';
      }
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'fileUrl') {
      void this.loadFromUrl();
    }

    // Handle auto-refresh toggle
    if (propertyPath === 'enableAutoRefresh') {
      if (newValue && this.properties.selectedMarkdownFile) {
        this.sharePointService.subscribeToFileChanges(this.properties.selectedMarkdownFile, () => {
          void this.checkAndReloadFile(this.properties.selectedMarkdownFile);
        });
      } else {
        this.sharePointService.unsubscribeFromFileChanges();
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
              groupName: 'Content Source',
              groupFields: [
                PropertyPaneDropdown('contentSource', {
                  label: 'Content Source',
                  options: [
                    { key: 'manual', text: 'Manual Entry' },
                    { key: 'browser', text: 'Library Browser' },
                    { key: 'url', text: 'Direct URL' }
                  ],
                  selectedKey: this.properties.contentSource
                }),
                ...(this.properties.contentSource === 'manual' ? [
                  PropertyPaneTextField('markdownContent', {
                    label: 'Markdown Content',
                    multiline: true,
                    rows: 15,
                    description: 'Enter your markdown content here'
                  })
                ] : []),
                ...(this.properties.contentSource === 'url' ? [
                  PropertyPaneTextField('fileUrl', {
                    label: 'File URL',
                    description: 'Enter the absolute URL to a markdown file (e.g., https://example.com/file.md)',
                    placeholder: 'https://...'
                  })
                ] : [])
              ]
            },
            ...(this.properties.contentSource === 'browser' ? [{
              groupName: 'Library Browser',
              groupFields: [
                PropertyPaneDropdown('selectedLibrary', {
                  label: 'Document Library',
                  options: this.fileSourceManager.libraryOptions,
                  selectedKey: this.properties.selectedLibrary
                }),
                PropertyPaneDropdown('selectedFolder', {
                  label: 'Folder',
                  options: this.fileSourceManager.folderOptions,
                  selectedKey: this.properties.selectedFolder,
                  disabled: !this.properties.selectedLibrary
                }),
                PropertyPaneDropdown('selectedMarkdownFile', {
                  label: 'Markdown File',
                  options: this.fileSourceManager.fileOptions,
                  selectedKey: this.properties.selectedMarkdownFile,
                  disabled: !this.properties.selectedLibrary
                })
              ]
            }] : []),
            {
              groupName: 'SharePoint Options',
              groupFields: [
                PropertyPaneToggle('enableAutoRefresh', {
                  label: 'Auto-refresh when file changes',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('enableVersionHistory', {
                  label: 'Show version history button',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            {
              groupName: 'Rendering Options',
              groupFields: [
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
                PropertyPaneDropdown('theme', {
                  label: 'Theme',
                  options: [
                    { key: 'light', text: 'Light' },
                    { key: 'dark', text: 'Dark' }
                  ],
                  selectedKey: this.properties.theme
                })
              ]
            }
          ]
        }
      ]
    };
  }
}