// Monaco Editor TypeScript declarations
declare global {
  interface Window {
    monaco: any;
    require: any;
  }
}

export class MonacoLoader {
  private static isLoaded: boolean = false;
  private static isLoading: boolean = false;
  private static loadPromise: Promise<any> | null = null;

  public static async loadMonaco(): Promise<any> {
    // Return existing promise if already loading
    if (this.isLoading && this.loadPromise) {
      return this.loadPromise;
    }

    // Return monaco if already loaded
    if (this.isLoaded && window.monaco) {
      return Promise.resolve(window.monaco);
    }

    this.isLoading = true;
    this.loadPromise = this.loadMonacoScript();
    
    try {
      const monaco = await this.loadPromise;
      this.isLoaded = true;
      this.isLoading = false;
      return monaco;
    } catch (error) {
      this.isLoading = false;
      this.loadPromise = null;
      throw error;
    }
  }

  private static loadMonacoScript(): Promise<any> {
    return new Promise((resolve, reject) => {
      // Check if Monaco is already available
      if (window.monaco) {
        resolve(window.monaco);
        return;
      }

      // Check if already loaded
      const existingScript = document.querySelector('script[src*="monaco-editor"]');
      if (existingScript) {
        // Wait for existing script to load
        existingScript.addEventListener('load', () => {
          if (window.monaco) {
            resolve(window.monaco);
          } else {
            reject(new Error('Monaco failed to load from existing script'));
          }
        });
        return;
      }

      console.log('📝 MonacoLoader: Loading Monaco Editor from CDN...');

      // Create loader script for Monaco Editor
      const loaderScript = document.createElement('script');
      loaderScript.src = 'https://cdn.jsdelivr.net/npm/monaco-editor@0.44.0/min/vs/loader.js';
      loaderScript.async = true;
      
      loaderScript.onload = () => {
        console.log('📝 MonacoLoader: Loader script loaded, configuring require...');
        
        // Configure require.js for Monaco
        window.require.config({
          paths: {
            'vs': 'https://cdn.jsdelivr.net/npm/monaco-editor@0.44.0/min/vs'
          }
        });

        // Load Monaco Editor
        window.require(['vs/editor/editor.main'], () => {
          console.log('✅ MonacoLoader: Monaco Editor loaded successfully');
          console.log('📝 MonacoLoader: Monaco version:', window.monaco?.editor?.VERSION || 'unknown');
          resolve(window.monaco);
        }, (error: any) => {
          console.error('❌ MonacoLoader: Failed to load Monaco Editor:', error);
          reject(new Error(`Failed to load Monaco Editor: ${error}`));
        });
      };

      loaderScript.onerror = () => {
        console.error('❌ MonacoLoader: Failed to load Monaco loader script');
        reject(new Error('Failed to load Monaco loader script'));
      };

      document.head.appendChild(loaderScript);
    });
  }

  public static createEditor(container: HTMLElement, options: any): Promise<any> {
    return this.loadMonaco().then(monaco => {
      console.log('📝 MonacoLoader: Creating editor instance...');
      const editor = monaco.editor.create(container, {
        language: 'markdown',
        theme: 'vs-light',
        automaticLayout: true,
        wordWrap: 'on',
        lineNumbers: 'on',
        minimap: { enabled: false },
        scrollBeyondLastLine: false,
        fontFamily: 'Monaco, Consolas, "Courier New", monospace',
        fontSize: 14,
        lineHeight: 1.6,
        renderLineHighlight: 'line',
        selectOnLineNumbers: true,
        roundedSelection: false,
        readOnly: false,
        cursorStyle: 'line',
        ...options
      });

      // Add markdown language features
      this.configureMarkdownLanguage(monaco);
      
      console.log('✅ MonacoLoader: Editor created successfully');
      return editor;
    });
  }

  private static configureMarkdownLanguage(monaco: any): void {
    // Configure markdown language features
    monaco.languages.setLanguageConfiguration('markdown', {
      comments: {
        blockComment: ['<!--', '-->']
      },
      brackets: [
        ['{', '}'],
        ['[', ']'],
        ['(', ')']
      ],
      autoClosingPairs: [
        { open: '{', close: '}' },
        { open: '[', close: ']' },
        { open: '(', close: ')' },
        { open: '"', close: '"' },
        { open: "'", close: "'" },
        { open: '`', close: '`' },
        { open: '**', close: '**' },
        { open: '*', close: '*' }
      ],
      surroundingPairs: [
        { open: '{', close: '}' },
        { open: '[', close: ']' },
        { open: '(', close: ')' },
        { open: '"', close: '"' },
        { open: "'", close: "'" },
        { open: '`', close: '`' },
        { open: '**', close: '**' },
        { open: '*', close: '*' }
      ]
    });

    // Add completion provider for markdown
    monaco.languages.registerCompletionItemProvider('markdown', {
      provideCompletionItems: (model: any, position: any) => {
        const word = model.getWordUntilPosition(position);
        const range = {
          startLineNumber: position.lineNumber,
          endLineNumber: position.lineNumber,
          startColumn: word.startColumn,
          endColumn: word.endColumn
        };

        return {
          suggestions: [
            {
              label: 'mermaid',
              kind: monaco.languages.CompletionItemKind.Snippet,
              insertText: '```mermaid\ngraph TD\n    A[Start] --> B[End]\n```',
              insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
              documentation: 'Insert Mermaid diagram',
              range: range
            },
            {
              label: 'code',
              kind: monaco.languages.CompletionItemKind.Snippet,
              insertText: '```${1:language}\n${2:code}\n```',
              insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
              documentation: 'Insert code block',
              range: range
            },
            {
              label: 'table',
              kind: monaco.languages.CompletionItemKind.Snippet,
              insertText: '| Header 1 | Header 2 |\n|----------|----------|\n| Cell 1   | Cell 2   |',
              insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
              documentation: 'Insert table',
              range: range
            },
            {
              label: 'math',
              kind: monaco.languages.CompletionItemKind.Snippet,
              insertText: '$$\n${1:equation}\n$$',
              insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
              documentation: 'Insert math block',
              range: range
            }
          ]
        };
      }
    });
  }

  public static setTheme(theme: string): void {
    if (window.monaco) {
      const monacoTheme = theme === 'dark' ? 'vs-dark' : 'vs-light';
      window.monaco.editor.setTheme(monacoTheme);
    }
  }
}