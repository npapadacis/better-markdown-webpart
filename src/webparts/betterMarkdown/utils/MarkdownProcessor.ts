import { escape } from '@microsoft/sp-lodash-subset';

// Mermaid is loaded from CDN as external
declare const mermaid: any;

// Use require for better compatibility with SPFx
const MarkdownIt = require('markdown-it');
const markdownItAttrs = require('markdown-it-attrs');
const markdownItFootnote = require('markdown-it-footnote');
const markdownItEmoji = require('markdown-it-emoji');
const markdownItAnchor = require('markdown-it-anchor');
const markdownItTOC = require('markdown-it-table-of-contents');
const markdownItAbbr = require('markdown-it-abbr');
const markdownItMultimdTable = require('markdown-it-multimd-table');
const hljs = require('highlight.js');
const katex = require('katex');

export interface IMarkdownProcessorOptions {
  enableSyntaxHighlighting: boolean;
  enableTOC: boolean;
  enableMath: boolean;
  enableMermaid: boolean;
  theme: string;
  styleClasses: {
    codeToolbar: string;
    toolbar: string;
    toolbarItem: string;
    codeLine: string;
    lineNumber: string;
    lineContent: string;
    tableOfContents: string;
    blockquote: string;
    mathBlock: string;
    mathError: string;
    mermaidContainer: string;
    mermaidError: string;
  };
}

export class MarkdownProcessor {
  private md: any;
  private options: IMarkdownProcessorOptions;

  constructor(options: IMarkdownProcessorOptions) {
    this.options = options;
    this.initializeMarkdownIt();
  }

  private initializeMarkdownIt(): void {
    this.md = new MarkdownIt({
      html: false,
      linkify: true,
      typographer: true,
      breaks: false,
      highlight: (str: string, lang: string) => {
        if (this.options.enableSyntaxHighlighting && lang && hljs.getLanguage(lang)) {
          try {
            const highlighted = hljs.highlight(str, { language: lang, ignoreIllegals: true }).value;
            return `<div class="${this.options.styleClasses.codeToolbar}"><pre class="prismjs language-${lang}"><code class="language-${lang}">${this.addLineNumbers(highlighted)}</code></pre><div class="${this.options.styleClasses.toolbar}"><div class="${this.options.styleClasses.toolbarItem}"><button>Copy</button></div></div></div>`;
          } catch (e) {
            console.error('Highlight error:', e);
          }
        }
        return `<div class="${this.options.styleClasses.codeToolbar}"><pre class="prismjs"><code>${this.addLineNumbers(this.md.utils.escapeHtml(str))}</code></pre><div class="${this.options.styleClasses.toolbar}"><div class="${this.options.styleClasses.toolbarItem}"><button>Copy</button></div></div></div>`;
      }
    });

    this.setupPlugins();
    this.setupCustomRenderers();
    
    if (this.options.enableMath) {
      this.addMathSupport();
    }

    if (this.options.enableMermaid) {
      this.addMermaidSupport();
    }

    this.addTabSupport();
  }

  private setupPlugins(): void {
    // Add extensions - handle both default and named exports
    try {
      const attrsPlugin = markdownItAttrs.default || markdownItAttrs;
      this.md.use(attrsPlugin, {
        leftDelimiter: '{',
        rightDelimiter: '}',
        allowedAttributes: ['id', 'class']
      });
    } catch (e) {
      console.error('Error loading markdown-it-attrs:', e);
    }

    try {
      const footnotePlugin = markdownItFootnote.default || markdownItFootnote;
      this.md.use(footnotePlugin);
    } catch (e) {
      console.error('Error loading markdown-it-footnote:', e);
    }

    try {
      const emojiPlugin = markdownItEmoji.default || markdownItEmoji;
      this.md.use(emojiPlugin);
    } catch (e) {
      console.error('Error loading markdown-it-emoji:', e);
    }

    try {
      const abbrPlugin = markdownItAbbr.default || markdownItAbbr;
      this.md.use(abbrPlugin);
    } catch (e) {
      console.error('Error loading markdown-it-abbr:', e);
    }

    // Enhanced table support with multimd-table
    try {
      const multimdTablePlugin = markdownItMultimdTable.default || markdownItMultimdTable;
      this.md.use(multimdTablePlugin, {
        multiline: true,    // Allow multiline content in cells
        rowspan: true,      // Enable row spanning with ^^
        headerless: false,  // Require table headers
        multibody: true,    // Allow multiple table bodies
        autolabel: true     // Auto-generate table IDs
      });
      console.log('🔧 MarkdownProcessor: Enhanced table support enabled with multimd-table');
    } catch (e) {
      console.error('Error loading markdown-it-multimd-table:', e);
    }

    if (this.options.enableTOC) {
      try {
        const anchorPlugin = markdownItAnchor.default || markdownItAnchor;
        this.md.use(anchorPlugin, {
          permalink: true,
          permalinkBefore: true,
          permalinkSymbol: '🔗'
        });
      } catch (e) {
        console.error('Error loading markdown-it-anchor:', e);
      }

      try {
        const tocPlugin = markdownItTOC.default || markdownItTOC;
        this.md.use(tocPlugin, {
          includeLevel: [1, 2],
          containerClass: this.options.styleClasses.tableOfContents,
          anchorLink: true,
          anchorLinkSymbol: '',
          anchorLinkBefore: false
        });
      } catch (e) {
        console.error('Error loading markdown-it-table-of-contents:', e);
      }
    }
  }

  private setupCustomRenderers(): void {
    // Custom renderer for Wiki.js-style blockquotes
    const defaultBlockquoteOpen = this.md.renderer.rules.blockquote_open || 
      function(tokens: any, idx: number, options: any, env: any, self: any) {
        return self.renderToken(tokens, idx, options);
      };
    
    this.md.renderer.rules.blockquote_open = (tokens: any, idx: number, options: any, env: any, self: any) => {
      const token = tokens[idx];
      const className = token.attrGet('class') || '';
      return `<blockquote class="${this.options.styleClasses.blockquote} ${className}">`;
    };

    // Custom renderer for lists
    const defaultBulletListOpen = this.md.renderer.rules.bullet_list_open || 
      function(tokens: any, idx: number, options: any, env: any, self: any) {
        return self.renderToken(tokens, idx, options);
      };
    
    this.md.renderer.rules.bullet_list_open = (tokens: any, idx: number, options: any, env: any, self: any) => {
      const token = tokens[idx];
      if (token.attrGet('class')) {
        const className = token.attrGet('class');
        return `<ul class="${className}">`;
      }
      return defaultBulletListOpen(tokens, idx, options, env, self);
    };
  }

  private addLineNumbers(code: string): string {
    const lines = code.split('\n');
    let lineNumberedCode = '';
    
    lines.forEach((line, index) => {
      const lineNumber = index + 1;
      lineNumberedCode += `<span class="${this.options.styleClasses.codeLine}">`;
      lineNumberedCode += `<span class="${this.options.styleClasses.lineNumber}">${lineNumber}</span>`;
      lineNumberedCode += `<span class="${this.options.styleClasses.lineContent}">${line}</span>`;
      lineNumberedCode += '</span>\n';
    });
    
    return lineNumberedCode.trim();
  }

  private addMathSupport(): void {
    // Inline math: $...$
    this.md.inline.ruler.before('escape', 'math_inline', (state: any, silent: boolean) => {
      if (state.src[state.pos] !== '$') {
        return false;
      }
      
      const start = state.pos + 1;
      let end = start;
      
      while (end < state.src.length && state.src[end] !== '$') {
        end++;
      }
      
      if (end >= state.src.length) {
        return false;
      }
      
      if (!silent) {
        const token = state.push('math_inline', 'math', 0);
        token.content = state.src.slice(start, end);
        token.markup = '$';
      }
      
      state.pos = end + 1;
      return true;
    });

    // Block math: $$...$$
    this.md.block.ruler.before('fence', 'math_block', (state: any, startLine: number, endLine: number, silent: boolean) => {
      let pos = state.bMarks[startLine] + state.tShift[startLine];
      let max = state.eMarks[startLine];
      
      if (pos + 2 > max) {
        return false;
      }
      if (state.src.slice(pos, pos + 2) !== '$$') {
        return false;
      }
      
      pos += 2;
      let firstLine = state.src.slice(pos, max);
      
      if (firstLine.trim().endsWith('$$')) {
        firstLine = firstLine.trim().slice(0, -2);
      }
      
      let nextLine = startLine;
      let lastLine = '';
      
      while (nextLine < endLine) {
        nextLine++;
        pos = state.bMarks[nextLine] + state.tShift[nextLine];
        max = state.eMarks[nextLine];
        
        if (pos < max && state.tShift[nextLine] < state.blkIndent) {
          break;
        }
        
        const line = state.src.slice(pos, max);
        if (line.trim().endsWith('$$')) {
          lastLine = line.trim().slice(0, -2);
          break;
        }
      }
      
      if (!silent) {
        const token = state.push('math_block', 'math', 0);
        token.content = (firstLine + '\n' + state.getLines(startLine + 1, nextLine, 0, false) + lastLine).trim();
        token.markup = '$$';
        token.map = [startLine, nextLine + 1];
      }
      
      state.line = nextLine + 1;
      return true;
    });

    // Renderers
    this.md.renderer.rules.math_inline = (tokens: any, idx: number) => {
      try {
        return katex.renderToString(tokens[idx].content, { throwOnError: false });
      } catch (e) {
        return `<span class="${this.options.styleClasses.mathError}">${escape(tokens[idx].content)}</span>`;
      }
    };

    this.md.renderer.rules.math_block = (tokens: any, idx: number) => {
      try {
        return `<div class="${this.options.styleClasses.mathBlock}">${katex.renderToString(tokens[idx].content, { 
          throwOnError: false,
          displayMode: true 
        })}</div>`;
      } catch (e) {
        return `<div class="${this.options.styleClasses.mathError}">${escape(tokens[idx].content)}</div>`;
      }
    };
  }

  private addMermaidSupport(): void {
    const defaultFenceRenderer = this.md.renderer.rules.fence || 
      function(tokens: any, idx: number, options: any, env: any, self: any) {
        return self.renderToken(tokens, idx, options);
      };

    this.md.renderer.rules.fence = (tokens: any, idx: number, options: any, env: any, self: any) => {
      const token = tokens[idx];
      const info = token.info.trim();
      
      if (info === 'mermaid') {
        const code = token.content;
        const id = 'mermaid-' + Math.random().toString(36).substr(2, 9);
        
        console.log('📝 MarkdownProcessor: Creating mermaid element', { id, codeLength: code.length });
        
        // Always create the mermaid element, regardless of library availability
        // The renderer will handle the actual rendering when mermaid is ready
        return `<div class="${this.options.styleClasses.mermaidContainer}"><pre class="mermaid" id="${id}">${escape(code)}</pre></div>`;
      }
      
      return defaultFenceRenderer(tokens, idx, options, env, self);
    };
  }

  private addTabSupport(): void {
    const defaultHeadingOpen = this.md.renderer.rules.heading_open || 
      function(tokens: any, idx: number, options: any, env: any, self: any) {
        return self.renderToken(tokens, idx, options);
      };

    let inTabset = false;
    let tabsetLevel = 0;

    this.md.renderer.rules.heading_open = (tokens: any, idx: number, options: any, env: any, self: any) => {
      const token = tokens[idx];
      const nextToken = tokens[idx + 1];
      
      if (nextToken && nextToken.content) {
        const className = token.attrGet('class');
        if (className && className.includes('tabset')) {
          inTabset = true;
          tabsetLevel = parseInt(token.tag.substring(1));
          return '<div class="tabset"><div class="tabs">';
        }
        
        if (inTabset) {
          const currentLevel = parseInt(token.tag.substring(1));
          if (currentLevel === tabsetLevel + 1) {
            return `<input type="radio" name="tabset-${tabsetLevel}" id="tab-${idx}"><label for="tab-${idx}">`;
          } else if (currentLevel <= tabsetLevel) {
            inTabset = false;
          }
        }
      }
      
      return defaultHeadingOpen(tokens, idx, options, env, self);
    };
  }

  public render(content: string): string {
    try {
      return this.md.render(content);
    } catch (e) {
      const errorMsg = (e as Error).message || 'Unknown error';
      return `<div class="error">Error rendering markdown: ${errorMsg}</div>`;
    }
  }

  public extractTOC(html: string): { tocHtml: string; mainHtml: string } {
    let tocHtml = '';
    let mainHtml = html;
    
    if (this.options.enableTOC && html.includes(`class="${this.options.styleClasses.tableOfContents}"`)) {
      const tocMatch = html.match(new RegExp(`<div class="${this.options.styleClasses.tableOfContents}"[^>]*>.*?</div>`, 's'));
      if (tocMatch) {
        tocHtml = tocMatch[0];
        mainHtml = html.replace(tocMatch[0], ''); // Remove TOC from main content
      }
    }

    return { tocHtml, mainHtml };
  }

  public updateOptions(newOptions: Partial<IMarkdownProcessorOptions>): void {
    this.options = { ...this.options, ...newOptions };
    this.initializeMarkdownIt();
  }
}