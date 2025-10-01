export class CodeBlockEnhancer {
  private styleClasses: {
    codeToolbar: string;
    lineContent: string;
  };

  constructor(styleClasses: { codeToolbar: string; lineContent: string }) {
    this.styleClasses = styleClasses;
  }

  public addCopyButtonFunctionality(containerElement: HTMLElement): void {
    console.log('📋 CodeBlockEnhancer: Adding copy functionality to:', containerElement);
    const codeToolbars = containerElement.querySelectorAll('.' + this.styleClasses.codeToolbar);
    console.log('📋 CodeBlockEnhancer: Found code toolbars:', codeToolbars.length);
    
    codeToolbars.forEach((toolbar: HTMLElement, index) => {
      console.log(`📋 CodeBlockEnhancer: Processing toolbar ${index}:`, toolbar);
      const copyButton = toolbar.querySelector('button');
      console.log(`📋 CodeBlockEnhancer: Found button in toolbar ${index}:`, copyButton);
      
      if (copyButton && !copyButton.dataset.enhanced) {
        console.log(`📋 CodeBlockEnhancer: Enhancing button ${index}`);
        
        // Mark as enhanced to avoid duplicate event listeners
        copyButton.dataset.enhanced = 'true';
        
        copyButton.addEventListener('click', (event) => {
          console.log('📋 CodeBlockEnhancer: Copy button clicked');
          event.preventDefault();
          event.stopPropagation();
          
          const codeElement = toolbar.querySelector('code');
          console.log('📋 CodeBlockEnhancer: Found code element:', codeElement);
          
          if (codeElement) {
            const textContent = this.extractCodeText(codeElement);
            console.log('📋 CodeBlockEnhancer: Extracted text:', textContent.substring(0, 100) + '...');
            this.copyToClipboard(textContent, copyButton);
          } else {
            console.error('📋 CodeBlockEnhancer: No code element found in toolbar');
          }
        });
        
        console.log(`📋 CodeBlockEnhancer: Successfully enhanced button ${index}`);
      } else if (copyButton?.dataset.enhanced) {
        console.log(`📋 CodeBlockEnhancer: Button ${index} already enhanced, skipping`);
      } else {
        console.log(`📋 CodeBlockEnhancer: No button found in toolbar ${index}`);
      }
    });
  }

  private extractCodeText(codeElement: HTMLElement): string {
    // Extract text content without line numbers
    const lineContents = codeElement.querySelectorAll('.' + this.styleClasses.lineContent);
    let textContent = '';
    
    if (lineContents.length > 0) {
      lineContents.forEach(line => {
        textContent += line.textContent + '\n';
      });
      textContent = textContent.trim();
    } else {
      // Fallback: get all text but remove line numbers
      const allText = codeElement.textContent || '';
      const lines = allText.split('\n');
      textContent = lines.map(line => {
        // Remove line numbers (assuming they're at the start and followed by spaces/tabs)
        return line.replace(/^\s*\d+\s*/, '');
      }).join('\n').trim();
    }
    
    return textContent;
  }

  private copyToClipboard(text: string, button: HTMLButtonElement): void {
    console.log('📋 CodeBlockEnhancer: Attempting to copy to clipboard:', text.length, 'characters');
    
    // Modern clipboard API
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(() => {
        console.log('📋 CodeBlockEnhancer: Successfully copied using navigator.clipboard');
        this.showCopyFeedback(button, 'Copied!', true);
      }).catch((error) => {
        console.error('📋 CodeBlockEnhancer: navigator.clipboard failed:', error);
        this.fallbackCopyToClipboard(text, button);
      });
    } else {
      // Fallback for older browsers or non-secure contexts
      console.log('📋 CodeBlockEnhancer: Using fallback copy method');
      this.fallbackCopyToClipboard(text, button);
    }
  }

  private fallbackCopyToClipboard(text: string, button: HTMLButtonElement): void {
    try {
      const textArea = document.createElement('textarea');
      textArea.value = text;
      textArea.style.position = 'fixed';
      textArea.style.left = '-999999px';
      textArea.style.top = '-999999px';
      document.body.appendChild(textArea);
      textArea.focus();
      textArea.select();
      
      const successful = document.execCommand('copy');
      document.body.removeChild(textArea);
      
      if (successful) {
        console.log('📋 CodeBlockEnhancer: Successfully copied using fallback method');
        this.showCopyFeedback(button, 'Copied!', true);
      } else {
        console.error('📋 CodeBlockEnhancer: Fallback copy method failed');
        this.showCopyFeedback(button, 'Error', false);
      }
    } catch (error) {
      console.error('📋 CodeBlockEnhancer: Fallback copy method threw error:', error);
      this.showCopyFeedback(button, 'Error', false);
    }
  }

  private showCopyFeedback(button: HTMLButtonElement, message: string, success: boolean): void {
    const originalText = button.textContent;
    button.textContent = message;
    button.style.backgroundColor = success ? '#28a745' : '#dc3545';
    
    setTimeout(() => {
      button.textContent = originalText;
      button.style.backgroundColor = '';
    }, 2000);
  }

  public removeEnhancedMarkers(containerElement: HTMLElement): void {
    const buttons = containerElement.querySelectorAll('button[data-enhanced]');
    buttons.forEach(button => {
      delete (button as HTMLButtonElement).dataset.enhanced;
    });
  }
}