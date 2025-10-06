import { SharePointService, IVersionInfo } from './SharePointService';

export interface IFileVersionManagerOptions {
  styles: any;
  sharePointService: SharePointService;
  onSave: (content: string) => Promise<boolean>;
  onVersionRestored: () => void;
}

export class FileVersionManager {
  private options: IFileVersionManagerOptions;

  constructor(options: IFileVersionManagerOptions) {
    this.options = options;
  }

  /**
   * Show version history modal for a SharePoint file
   */
  public async showVersionHistory(fileUrl: string): Promise<void> {
    try {
      if (!fileUrl) {
        alert('No SharePoint file selected');
        return;
      }

      // Get version history
      const versions = await this.options.sharePointService.getFileVersions(fileUrl);

      if (versions.length === 0) {
        alert('No version history available for this file');
        return;
      }

      // Create modal
      const modal = this.createVersionModal(versions);
      document.body.appendChild(modal);

      // Set up event handlers
      this.setupModalHandlers(modal, versions);

    } catch (error) {
      console.error('Error showing version history:', error);
      alert(`Failed to load version history: ${error.message}`);
    }
  }

  private createVersionModal(versions: IVersionInfo[]): HTMLDivElement {
    const modal = document.createElement('div');
    modal.className = this.options.styles.versionModal;
    modal.innerHTML = `
      <div class="${this.options.styles.versionModalContent}">
        <div class="${this.options.styles.versionModalHeader}">
          <h2>Version History</h2>
          <button class="${this.options.styles.closeButton}" id="closeVersionModal">×</button>
        </div>
        <div class="${this.options.styles.versionList}">
          ${versions.map(version => `
            <div class="${this.options.styles.versionItem} ${version.isCurrentVersion ? this.options.styles.currentVersion : ''}" data-version-url="${version.url}">
              <div class="${this.options.styles.versionLabel}">
                Version ${version.versionLabel}${version.isCurrentVersion ? ' (Current)' : ''}
              </div>
              <div class="${this.options.styles.versionInfo}">
                <span class="${this.options.styles.versionDate}">
                  📅 ${new Date(version.created).toLocaleString()}
                </span>
                <span class="${this.options.styles.versionAuthor}">
                  👤 ${version.createdBy}
                </span>
              </div>
              ${!version.isCurrentVersion ? `
                <div class="${this.options.styles.versionActions}">
                  <button class="restoreVersion" data-version-url="${version.url}">Restore This Version</button>
                  <button class="viewButton viewVersion" data-version-url="${version.url}">View Content</button>
                </div>
              ` : ''}
            </div>
          `).join('')}
        </div>
      </div>
    `;
    return modal;
  }

  private setupModalHandlers(modal: HTMLDivElement, versions: IVersionInfo[]): void {
    // Close modal on click outside or close button
    modal.addEventListener('click', (e) => {
      if (e.target === modal || (e.target as HTMLElement).id === 'closeVersionModal') {
        document.body.removeChild(modal);
      }
    });

    // Restore version handlers
    modal.querySelectorAll('.restoreVersion').forEach(button => {
      button.addEventListener('click', async (e) => {
        await this.handleRestoreVersion(e, modal);
      });
    });

    // View version handlers
    modal.querySelectorAll('.viewVersion').forEach(button => {
      button.addEventListener('click', async (e) => {
        await this.handleViewVersion(e);
      });
    });
  }

  private async handleRestoreVersion(e: Event, modal: HTMLDivElement): Promise<void> {
    const versionUrl = (e.target as HTMLElement).getAttribute('data-version-url');
    if (versionUrl && confirm('Are you sure you want to restore this version? This will create a new version with the old content.')) {
      try {
        // Fetch the version content
        const versionContent = await fetch(versionUrl).then(r => r.text());

        // Save as new version
        const success = await this.options.onSave(versionContent);

        if (success) {
          alert('Version restored successfully!');
          document.body.removeChild(modal);
          this.options.onVersionRestored();
        }
      } catch (error) {
        console.error('Error restoring version:', error);
        alert('Failed to restore version');
      }
    }
  }

  private async handleViewVersion(e: Event): Promise<void> {
    const versionUrl = (e.target as HTMLElement).getAttribute('data-version-url');
    if (versionUrl) {
      try {
        const versionContent = await fetch(versionUrl).then(r => r.text());

        // Show in a new window
        const viewWindow = window.open('', '_blank', 'width=800,height=600');
        if (viewWindow) {
          viewWindow.document.write(`
            <html>
              <head><title>Version Content</title></head>
              <body style="padding: 20px; font-family: monospace; white-space: pre-wrap;">
                ${versionContent.replace(/</g, '&lt;').replace(/>/g, '&gt;')}
              </body>
            </html>
          `);
          viewWindow.document.close();
        }
      } catch (error) {
        console.error('Error viewing version:', error);
        alert('Failed to load version content');
      }
    }
  }
}
