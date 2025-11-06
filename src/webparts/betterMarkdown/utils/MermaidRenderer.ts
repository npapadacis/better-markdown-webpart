declare const mermaid: any;

export class MermaidRenderer {
  private isInitialized: boolean = false;
  private theme: string = 'default';

  public get initialized(): boolean {
    return this.isInitialized;
  }

  public async initialize(theme: string = 'default'): Promise<void> {
    console.log('🔍 MermaidRenderer: Starting initialization...');
    console.log('🔍 Initial mermaid check:', typeof mermaid);
    console.log('🔍 Window object has mermaid:', 'mermaid' in window);
    console.log('🔍 Global mermaid:', (window as any).mermaid);
    
    this.theme = theme;
    
    // Check if mermaid is already available on window
    const globalMermaid = (window as any).mermaid;
    if (globalMermaid) {
      console.log('✅ Mermaid found on window object');
      try {
        await this.initializeMermaid(globalMermaid, theme);
        return;
      } catch (e) {
        console.error('❌ Failed to initialize window.mermaid:', e);
      }
    }
    
    // Check if mermaid is available via the external declaration
    if (typeof mermaid !== 'undefined') {
      console.log('✅ Mermaid found via external declaration');
      try {
        await this.initializeMermaid(mermaid, theme);
        return;
      } catch (e) {
        console.error('❌ Failed to initialize external mermaid:', e);
      }
    }
    
    console.log('⏳ Mermaid not immediately available, attempting to load script...');
    
    // Try to load the script manually
    try {
      await this.loadMermaidScript();
      console.log('✅ Script loaded, checking for mermaid again...');
      
      // Wait and retry
      const mermaidInstance = await this.waitForMermaid();
      if (mermaidInstance) {
        await this.initializeMermaid(mermaidInstance, theme);
      } else {
        console.error('❌ Mermaid still not available after loading script');
      }
    } catch (e) {
      console.error('❌ Failed to load mermaid script:', e);
    }
  }

  private async initializeMermaid(mermaidInstance: any, theme: string): Promise<void> {
    console.log('🔧 Initializing mermaid with theme:', theme);
    console.log('🔧 Mermaid instance:', mermaidInstance);
    console.log('🔧 Mermaid version:', mermaidInstance.version || 'unknown');
    
    try {
      await mermaidInstance.initialize({
        startOnLoad: false,
        theme: theme === 'dark' ? 'dark' : 'default',
        securityLevel: 'loose',
        // Configuration to handle long text labels
        flowchart: {
          htmlLabels: true,
          curve: 'basis',
          padding: 15,
          nodeSpacing: 50,
          rankSpacing: 50,
          wrappingWidth: 200
        },
        // General text handling
        wrap: true,
        fontSize: 16,
        fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
        // Layout improvements
        gantt: {
          fontSize: 14,
          fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif'
        },
        sequence: {
          fontSize: 14,
          fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
          wrap: true,
          width: 150
        }
      });
      this.isInitialized = true;
      console.log('✅ Mermaid initialized successfully');
    } catch (e) {
      console.error('❌ Error during mermaid initialization:', e);
      throw e;
    }
  }

  private async waitForMermaid(maxAttempts: number = 50): Promise<any> {
    console.log('⏳ Waiting for mermaid to become available...');
    
    for (let attempts = 0; attempts < maxAttempts; attempts++) {
      // Check multiple ways mermaid might be available
      const checks = [
        () => typeof mermaid !== 'undefined' ? mermaid : null,
        () => (window as any).mermaid || null,
        () => (window as any).Mermaid || null
      ];
      
      for (const check of checks) {
        const result = check();
        if (result) {
          console.log(`✅ Mermaid found after ${attempts} attempts via check:`, check.toString());
          return result;
        }
      }
      
      if (attempts % 10 === 0) {
        console.log(`⏳ Still waiting... attempt ${attempts}/${maxAttempts}`);
      }
      
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    
    console.error('❌ Mermaid not found after maximum attempts');
    return null;
  }

  private loadMermaidScript(): Promise<void> {
    return new Promise((resolve, reject) => {
      // Check if script is already loaded
      const existingScript = document.querySelector('script[src*="mermaid"]');
      if (existingScript) {
        resolve();
        return;
      }

      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/mermaid@11.12.0/dist/mermaid.min.js';
      script.async = true;
      script.onload = () => {
        console.log('Mermaid script loaded successfully');
        resolve();
      };
      script.onerror = () => {
        console.error('Failed to load Mermaid script');
        reject(new Error('Failed to load Mermaid script'));
      };
      
      document.head.appendChild(script);
    });
  }

  public async renderDiagrams(containerElement: HTMLElement, errorStyleClass: string): Promise<void> {
    console.log('🎨 MermaidRenderer: Starting diagram rendering...');
    console.log('🎨 Is initialized:', this.isInitialized);
    console.log('🎨 Mermaid available:', typeof mermaid !== 'undefined');
    console.log('🎨 Window.mermaid available:', !!(window as any).mermaid);

    // First, let's see what's in the container
    console.log('🎨 Container innerHTML length:', containerElement.innerHTML.length);
    console.log('🎨 Container has .mermaid class elements:', containerElement.innerHTML.includes('class="mermaid"'));

    const mermaidElements = containerElement.querySelectorAll('.mermaid');
    console.log('🎨 Found mermaid elements:', mermaidElements.length);
    
    // Log details about each element found
    mermaidElements.forEach((element, index) => {
      console.log(`🎨 Mermaid element ${index}:`, {
        id: element.id,
        className: element.className,
        textContent: element.textContent?.substring(0, 50) + '...'
      });
    });
    
    if (mermaidElements.length === 0) {
      console.log('⚠️ No mermaid elements found to render');
      console.log('🔍 Looking for any mermaid-related content...');
      const allPreElements = containerElement.querySelectorAll('pre');
      console.log('🔍 Total pre elements found:', allPreElements.length);
      allPreElements.forEach((pre, index) => {
        console.log(`🔍 Pre element ${index}:`, {
          className: pre.className,
          textContent: pre.textContent?.substring(0, 30) + '...'
        });
      });
      return;
    }

    // Get the mermaid instance
    let mermaidInstance = null;
    if (typeof mermaid !== 'undefined') {
      mermaidInstance = mermaid;
      console.log('✅ Using external mermaid declaration');
    } else if ((window as any).mermaid) {
      mermaidInstance = (window as any).mermaid;
      console.log('✅ Using window.mermaid');
    } else {
      console.error('❌ No mermaid instance available for rendering');
      // Show error message in all mermaid elements
      mermaidElements.forEach(element => {
        element.innerHTML = `<div class="${errorStyleClass}">Mermaid library not loaded. Please check console for details.</div>`;
      });
      return;
    }

    console.log('🎨 Mermaid instance found:', !!mermaidInstance);
    console.log('🎨 Mermaid methods available:', Object.keys(mermaidInstance).join(', '));
    
    for (let i = 0; i < mermaidElements.length; i++) {
      const element = mermaidElements[i] as HTMLElement;
      const code = element.textContent || '';

      // Skip elements that have already been rendered (contain SVG)
      if (code.includes('<svg') || code.includes('#mermaid-') && code.includes('-svg{')) {
        console.log(`⏭️ Skipping diagram ${i + 1}: already rendered`);
        continue;
      }

      const id = element.id || `mermaid-${i}-${Date.now()}`;
      console.log(`🎨 Rendering diagram ${i + 1}:`, { id, codeLength: code.length });

      try {
        // For newer versions of Mermaid, use the render method
        if (mermaidInstance.render) {
          const result = await mermaidInstance.render(id + '-svg', code);
          element.innerHTML = result.svg;
          console.log(`✅ Successfully rendered diagram ${i + 1}`);
        } else if (mermaidInstance.mermaidAPI && mermaidInstance.mermaidAPI.render) {
          // Fallback for older versions
          const result = await mermaidInstance.mermaidAPI.render(id + '-svg', code);
          element.innerHTML = result;
          console.log(`✅ Successfully rendered diagram ${i + 1} (legacy API)`);
        } else {
          throw new Error('No suitable render method found on mermaid instance');
        }
      } catch (e) {
        const errorMsg = (e as Error).message || 'Unknown error';
        console.error(`❌ Error rendering diagram ${i + 1}:`, e);
        element.innerHTML = `<div class="${errorStyleClass}">Error rendering diagram: ${errorMsg}</div>`;
      }
    }
  }

  public updateTheme(newTheme: string): void {
    console.log('🎨 MermaidRenderer: Updating theme from', this.theme, 'to', newTheme);
    
    if (this.theme !== newTheme && this.isInitialized) {
      this.theme = newTheme;
      
      const mermaidInstance = (window as any).mermaid || (typeof mermaid !== 'undefined' ? mermaid : null);
      if (mermaidInstance) {
        try {
          mermaidInstance.initialize({
            startOnLoad: false,
            theme: newTheme === 'dark' ? 'dark' : 'default',
            securityLevel: 'loose',
            // Configuration to handle long text labels
            flowchart: {
              htmlLabels: true,
              curve: 'basis',
              padding: 15,
              nodeSpacing: 50,
              rankSpacing: 50,
              wrappingWidth: 200
            },
            // General text handling
            wrap: true,
            fontSize: 16,
            fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
            // Layout improvements
            gantt: {
              fontSize: 14,
              fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif'
            },
            sequence: {
              fontSize: 14,
              fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
              wrap: true,
              width: 150
            }
          });
          console.log('✅ MermaidRenderer: Theme updated successfully');
        } catch (e) {
          console.error('❌ MermaidRenderer: Error updating theme:', e);
        }
      } else {
        console.log('⚠️ MermaidRenderer: Mermaid not available for theme update');
      }
    }
  }
}