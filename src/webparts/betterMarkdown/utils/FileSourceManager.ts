import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { SharePointService, IFileMetadata } from './SharePointService';

export interface IFileSourceManagerOptions {
  sharePointService: SharePointService;
  onContentLoaded: (content: string, metadata: IFileMetadata | null, lastModified?: string) => void;
  onAutoRefreshSetup: (fileUrl: string, callback: () => void) => void;
}

export class FileSourceManager {
  private options: IFileSourceManagerOptions;
  public libraryOptions: IPropertyPaneDropdownOption[] = [];
  public folderOptions: IPropertyPaneDropdownOption[] = [];
  public fileOptions: IPropertyPaneDropdownOption[] = [];

  constructor(options: IFileSourceManagerOptions) {
    this.options = options;
  }

  /**
   * Load all document libraries in the current site
   */
  public async loadLibraryOptions(selectedLibrary?: string): Promise<void> {
    const libraries = await this.options.sharePointService.getDocumentLibraries();
    this.libraryOptions = libraries.map(lib => ({
      key: lib.serverRelativeUrl,
      text: lib.title
    }));

    // If a library is selected, load folder and file options
    if (selectedLibrary) {
      await this.loadFolderOptions(selectedLibrary);
      await this.loadFileOptions(selectedLibrary, '');
    }
  }

  /**
   * Load folders from a library
   */
  public async loadFolderOptions(libraryUrl: string): Promise<void> {
    if (!libraryUrl) return;

    const folders = await this.options.sharePointService.getFolders(libraryUrl);
    this.folderOptions = [
      { key: '', text: '(Root)' },
      ...folders.map(folder => ({
        key: folder,
        text: folder
      }))
    ];
  }

  /**
   * Load markdown files from a library/folder
   */
  public async loadFileOptions(libraryUrl: string, folderPath: string): Promise<void> {
    if (!libraryUrl) return;

    const files = await this.options.sharePointService.getMarkdownFiles(
      libraryUrl,
      folderPath
    );

    this.fileOptions = files.map(file => ({
      key: file.serverRelativeUrl,
      text: file.name
    }));
  }

  /**
   * Load content from a selected SharePoint file
   */
  public async loadSelectedFile(
    fileUrl: string,
    enableAutoRefresh: boolean,
    autoRefreshCallback: () => void
  ): Promise<void> {
    if (!fileUrl) return;

    try {
      // Unsubscribe from previous file changes
      this.options.sharePointService.unsubscribeFromFileChanges();

      const content = await this.options.sharePointService.getFileContent(fileUrl);
      const metadata = await this.options.sharePointService.getFileMetadata(fileUrl);

      this.options.onContentLoaded(content, metadata, metadata?.timeLastModified);

      // Set up auto-refresh if enabled
      if (enableAutoRefresh) {
        console.log('📡 Setting up auto-refresh subscription for:', fileUrl);
        this.options.sharePointService.subscribeToFileChanges(fileUrl, autoRefreshCallback);
      }
    } catch (error) {
      console.error('Error loading selected file:', error);
      alert(`Failed to load file: ${error.message}`);
    }
  }

  /**
   * Load content from a URL
   */
  public async loadFromUrl(fileUrl: string): Promise<void> {
    if (!fileUrl) return;

    try {
      // Security check: Only allow text/plain or markdown content types
      const response = await fetch(fileUrl, {
        headers: {
          'Accept': 'text/plain, text/markdown, text/x-markdown, */*'
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      // Validate content type
      const contentType = response.headers.get('content-type') || '';
      const allowedTypes = ['text/plain', 'text/markdown', 'text/x-markdown', 'application/octet-stream'];
      const isAllowed = allowedTypes.some(type => contentType.toLowerCase().includes(type));

      if (!isAllowed && contentType) {
        console.warn(`⚠️ Content type "${contentType}" may not be markdown. Proceeding with caution.`);
      }

      const content = await response.text();

      // Basic security: Check for script tags or HTML
      if (content.includes('<script') || content.includes('javascript:')) {
        throw new Error('Security: File contains potentially dangerous content (script tags or javascript: URLs)');
      }

      const metadata: IFileMetadata = {
        name: fileUrl.split('/').pop() || 'Unknown',
        serverRelativeUrl: fileUrl,
        timeLastModified: new Date().toISOString(),
        author: 'External',
        authorEmail: '',
        length: content.length,
        listItemId: 0,
        uniqueId: ''
      };

      this.options.onContentLoaded(content, metadata);

    } catch (error) {
      console.error('Error loading from URL:', error);
      alert(`Failed to load from URL: ${error.message}`);
    }
  }

  /**
   * Unsubscribe from file change notifications
   */
  public dispose(): void {
    this.options.sharePointService.unsubscribeFromFileChanges();
  }
}
