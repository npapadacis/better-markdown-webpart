import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';

export interface IFileMetadata {
  name: string;
  serverRelativeUrl: string;
  timeLastModified: string;
  author: string;
  authorEmail: string;
  length: number;
  listItemId: number;
  uniqueId: string;
}

export interface IVersionInfo {
  versionLabel: string;
  created: string;
  createdBy: string;
  url: string;
  isCurrentVersion: boolean;
}

export interface ILibraryInfo {
  title: string;
  serverRelativeUrl: string;
  itemCount: number;
}

export class SharePointService {
  private context: WebPartContext;
  private sp: SPFI;
  private lastKnownModifiedTime: string | null = null;

  constructor(context: WebPartContext) {
    this.context = context;

    // Initialize PnP SP v4
    this.sp = spfi().using(SPFx(this.context));
  }

  /**
   * Get all document libraries in the current site
   */
  public async getDocumentLibraries(): Promise<ILibraryInfo[]> {
    try {
      const lists = await this.sp.web.lists
        .filter("BaseTemplate eq 101 and Hidden eq false") // 101 = Document Library
        .select("Title", "RootFolder/ServerRelativeUrl", "ItemCount")
        .expand("RootFolder")();

      return lists.map((list: any) => ({
        title: list.Title,
        serverRelativeUrl: list.RootFolder.ServerRelativeUrl,
        itemCount: list.ItemCount
      }));
    } catch (error) {
      console.error('Error getting document libraries:', error);
      return [];
    }
  }

  /**
   * Get markdown files from a specific library or folder
   */
  public async getMarkdownFiles(libraryUrl: string, folderPath?: string): Promise<IFileMetadata[]> {
    try {
      // Only append folder path if it's not empty
      const targetPath = (folderPath && folderPath.trim() !== '') ? `${libraryUrl}/${folderPath}` : libraryUrl;

      console.log('📁 SharePointService: Getting markdown files from:', targetPath);

      const files = await this.sp.web.getFolderByServerRelativePath(targetPath).files
        .select(
          "Name",
          "ServerRelativeUrl",
          "TimeLastModified",
          "Author/Title",
          "Author/EMail",
          "Length",
          "ListItemAllFields/Id",
          "UniqueId"
        )
        .expand("Author", "ListItemAllFields")();

      // Filter for markdown files client-side
      const markdownFiles = files.filter((file: any) => {
        const name = file.Name.toLowerCase();
        return name.endsWith('.md') || name.endsWith('.markdown') || name.endsWith('.txt');
      });

      const results = markdownFiles.map((file: any) => ({
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        timeLastModified: file.TimeLastModified,
        author: file.Author?.Title || 'Unknown',
        authorEmail: file.Author?.EMail || '',
        length: file.Length,
        listItemId: file.ListItemAllFields?.Id || 0,
        uniqueId: file.UniqueId
      }));

      console.log(`📁 SharePointService: Found ${results.length} markdown files out of ${files.length} total files`);
      return results;
    } catch (error) {
      console.error('Error getting markdown files:', error);
      return [];
    }
  }

  /**
   * Get file content
   */
  public async getFileContent(fileUrl: string): Promise<string> {
    try {
      const file = await this.sp.web.getFileByServerRelativePath(fileUrl).getText();
      return file;
    } catch (error) {
      console.error('Error getting file content:', error);
      throw error;
    }
  }

  /**
   * Get file metadata
   */
  public async getFileMetadata(fileUrl: string): Promise<IFileMetadata | null> {
    try {
      const file: any = await this.sp.web.getFileByServerRelativePath(fileUrl)
        .select(
          "Name",
          "ServerRelativeUrl",
          "TimeLastModified",
          "Author/Title",
          "Author/EMail",
          "Length",
          "ListItemAllFields/Id",
          "UniqueId"
        )
        .expand("Author", "ListItemAllFields")
        ();

      return {
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        timeLastModified: file.TimeLastModified,
        author: file.Author?.Title || 'Unknown',
        authorEmail: file.Author?.EMail || '',
        length: file.Length,
        listItemId: file.ListItemAllFields?.Id || 0,
        uniqueId: file.UniqueId
      };
    } catch (error) {
      console.error('Error getting file metadata:', error);
      return null;
    }
  }

  /**
   * Get file version history
   */
  public async getFileVersions(fileUrl: string): Promise<IVersionInfo[]> {
    try {
      const versions = await this.sp.web.getFileByServerRelativePath(fileUrl)
        .versions
        .select("VersionLabel", "Created", "CreatedBy/Title", "Url", "IsCurrentVersion")
        .expand("CreatedBy")
        ();

      return versions.map((version: any) => ({
        versionLabel: version.VersionLabel,
        created: version.Created,
        createdBy: version.CreatedBy?.Title || 'Unknown',
        url: version.Url,
        isCurrentVersion: version.IsCurrentVersion
      }));
    } catch (error) {
      console.error('Error getting file versions:', error);
      return [];
    }
  }

  /**
   * Save content to SharePoint file
   */
  public async saveFileContent(fileUrl: string, content: string): Promise<boolean> {
    try {
      await this.sp.web.getFileByServerRelativePath(fileUrl)
        .setContent(content);
      return true;
    } catch (error) {
      console.error('Error saving file content:', error);
      return false;
    }
  }

  /**
   * Create a new file in SharePoint
   */
  public async createFile(libraryUrl: string, fileName: string, content: string): Promise<string | null> {
    try {
      const file = await this.sp.web.getFolderByServerRelativePath(libraryUrl)
        .files.addUsingPath(fileName, content, { Overwrite: true });

      return file.ServerRelativeUrl;
    } catch (error) {
      console.error('Error creating file:', error);
      return null;
    }
  }

  /**
   * Check if file has been modified since last load
   */
  public async checkFileModified(fileUrl: string, lastModified: string): Promise<boolean> {
    try {
      const file = await this.sp.web.getFileByServerRelativePath(fileUrl)
        .select("TimeLastModified")
        ();

      const fileModified = new Date(file.TimeLastModified);
      const lastCheck = new Date(lastModified);

      return fileModified > lastCheck;
    } catch (error) {
      console.error('Error checking file modification:', error);
      return false;
    }
  }

  /**
   * Get folders in a library
   */
  public async getFolders(libraryUrl: string, parentFolder?: string): Promise<string[]> {
    try {
      const targetPath = parentFolder ? `${libraryUrl}/${parentFolder}` : libraryUrl;

      const folders = await this.sp.web.getFolderByServerRelativePath(targetPath)
        .folders
        .select("Name")
        .filter("Name ne 'Forms'") // Exclude Forms folder
        ();

      return folders.map((folder: any) => folder.Name);
    } catch (error) {
      console.error('Error getting folders:', error);
      return [];
    }
  }

  /**
   * Subscribe to file changes (for auto-refresh)
   */
  public async subscribeToFileChanges(fileUrl: string, callback: () => void): Promise<void> {
    // Initialize with current file timestamp
    try {
      const metadata = await this.getFileMetadata(fileUrl);
      if (metadata) {
        this.lastKnownModifiedTime = metadata.timeLastModified;
        console.log('📡 Auto-refresh: Starting subscription for', fileUrl, 'Last modified:', this.lastKnownModifiedTime);
      }
    } catch (error) {
      console.error('Error initializing file change subscription:', error);
    }

    // Poll for changes every 30 seconds
    const checkInterval = setInterval(async () => {
      try {
        console.log('📡 Auto-refresh: Checking for file changes...');
        const metadata = await this.getFileMetadata(fileUrl);
        if (metadata) {
          console.log('📡 Auto-refresh: Current timestamp:', metadata.timeLastModified);
          console.log('📡 Auto-refresh: Last known timestamp:', this.lastKnownModifiedTime);
          // Check if file was modified since last check
          if (this.lastKnownModifiedTime && metadata.timeLastModified !== this.lastKnownModifiedTime) {
            console.log('📡 Auto-refresh: File change detected!');
            console.log('   Previous:', this.lastKnownModifiedTime);
            console.log('   Current:', metadata.timeLastModified);
            this.lastKnownModifiedTime = metadata.timeLastModified;
            callback();
          } else {
            console.log('📡 Auto-refresh: No changes detected');
          }
        }
      } catch (error) {
        console.error('Error in file change subscription:', error);
        clearInterval(checkInterval);
      }
    }, 30000); // Check every 30 seconds

    // Store interval ID for cleanup
    (window as any).__fileChangeInterval = checkInterval;
  }

  /**
   * Unsubscribe from file changes
   */
  public unsubscribeFromFileChanges(): void {
    const intervalId = (window as any).__fileChangeInterval;
    if (intervalId) {
      clearInterval(intervalId);
      delete (window as any).__fileChangeInterval;
      this.lastKnownModifiedTime = null;
      console.log('📡 Auto-refresh: Unsubscribed from file changes');
    }
  }
}
