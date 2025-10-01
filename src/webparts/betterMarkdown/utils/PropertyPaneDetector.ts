export class PropertyPaneDetector {
  private observer: MutationObserver | null = null;
  private onStateChange: (isOpen: boolean) => void;
  private currentState: boolean = false;
  private debounceTimer: number | null = null;

  constructor(onStateChange: (isOpen: boolean) => void) {
    this.onStateChange = onStateChange;
  }

  public start(): void {
    // Initial check
    this.checkPropertyPaneState();

    // Set up observer to watch for property pane changes
    if (typeof MutationObserver !== 'undefined') {
      this.observer = new MutationObserver((mutations) => {
        // Check if any mutations might affect property pane visibility
        let shouldCheck = false;
        
        mutations.forEach(mutation => {
          if (mutation.type === 'childList') {
            // Check if any added/removed nodes might be property pane related
            const checkNodes = (nodes: NodeList) => {
              nodes.forEach(node => {
                if (node.nodeType === Node.ELEMENT_NODE) {
                  const element = node as Element;
                  const className = (element.className || '').toString();
                  const dataId = (element.getAttribute('data-automation-id') || '').toString();
                  if (className.toLowerCase().includes('property') || 
                      className.toLowerCase().includes('panel') ||
                      dataId.toLowerCase().includes('property')) {
                    shouldCheck = true;
                  }
                }
              });
            };
            
            if (mutation.addedNodes.length > 0) checkNodes(mutation.addedNodes);
            if (mutation.removedNodes.length > 0) checkNodes(mutation.removedNodes);
          } else if (mutation.type === 'attributes') {
            const target = mutation.target as Element;
            const className = (target.className || '').toString();
            const dataId = (target.getAttribute('data-automation-id') || '').toString();
            if (className.toLowerCase().includes('property') || 
                className.toLowerCase().includes('panel') ||
                dataId.toLowerCase().includes('property')) {
              shouldCheck = true;
            }
          }
        });
        
        if (shouldCheck) {
          // Debounce multiple rapid changes to prevent flickering
          if (this.debounceTimer) {
            clearTimeout(this.debounceTimer);
          }
          this.debounceTimer = window.setTimeout(() => {
            this.checkPropertyPaneState();
            this.debounceTimer = null;
          }, 100);
        }
      });

      // Observe changes to the document body
      this.observer.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['class', 'style', 'data-automation-id', 'aria-label']
      });
    }
  }

  public stop(): void {
    if (this.observer) {
      this.observer.disconnect();
      this.observer = null;
    }
    if (this.debounceTimer) {
      clearTimeout(this.debounceTimer);
      this.debounceTimer = null;
    }
  }

  private checkPropertyPaneState(): void {
    // Check multiple possible selectors for SPFx property pane
    const selectors = [
      '[data-automation-id="propertyPanePageCard"]',
      '.sp-PropertyPane:not([class*="Container"]):not(aside)',
      '[class*="propertyPane"]:not([class*="Container"]):not(aside)',
      '[class*="PropertyPane"]:not([class*="Container"]):not(aside)',
      '[data-automation-id*="propertyPane"]:not([data-automation-id*="Container"]):not(aside)',
      '[data-automation-id*="PropertyPane"]:not([data-automation-id*="Container"]):not(aside)',
      '.ms-Panel:not([class*="Container"]):not(aside)',
      '[class*="Panel"][class*="property"]:not([class*="Container"]):not(aside)',
      '[aria-label*="property" i][aria-label*="pane" i]:not([class*="Container"]):not(aside)'
    ];

    let isPropertyPaneOpen = false;
    let foundSelector = '';
    let foundElement: Element | null = null;

    console.log('🔍 PropertyPaneDetector: Checking all selectors...');

    for (const selector of selectors) {
      const elements = document.querySelectorAll(selector);
      console.log(`🔍 PropertyPaneDetector: Selector "${selector}" found ${elements.length} elements`);
      
      for (let index = 0; index < elements.length; index++) {
        const element = elements[index];
        const computedStyle = window.getComputedStyle(element);
        const rect = element.getBoundingClientRect();
        
        console.log(`🔍 PropertyPaneDetector: Element ${index} for "${selector}":`, {
          tagName: element.tagName,
          className: element.className,
          id: element.id,
          'data-automation-id': element.getAttribute('data-automation-id'),
          display: computedStyle.display,
          visibility: computedStyle.visibility,
          opacity: computedStyle.opacity,
          width: rect.width,
          height: rect.height,
          top: rect.top,
          left: rect.left
        });

        // Additional check to ensure it's actually visible and positioned like a property pane
        if (computedStyle.display !== 'none' && 
            computedStyle.visibility !== 'hidden' &&
            computedStyle.opacity !== '0') {
          // Check if element has actual dimensions
          if (rect.width > 0 && rect.height > 0) {
            // Additional check: property pane should be positioned on the right side of the screen
            const isOnRightSide = rect.left > (window.innerWidth * 0.6); // Property pane typically appears on right 40% of screen
            const hasReasonableSize = rect.width > 250 && rect.height > 300; // Property pane is typically at least 250px wide and 300px tall
            
            console.log(`🔍 PropertyPaneDetector: Element ${index} visibility checks:`, {
              isVisible: true,
              isOnRightSide,
              hasReasonableSize,
              position: `${rect.left}, ${rect.top}`,
              size: `${rect.width}x${rect.height}`
            });

            if (isOnRightSide && hasReasonableSize) {
              isPropertyPaneOpen = true;
              foundSelector = selector;
              foundElement = element;
              console.log(`🔍 PropertyPaneDetector: ✅ Property pane detected via "${selector}"`);
              break;
            } else {
              console.log(`🔍 PropertyPaneDetector: ❌ Element "${selector}" doesn't match property pane criteria`);
            }
          }
        }
      }
      
      if (isPropertyPaneOpen) break;
    }

    console.log('🔍 PropertyPaneDetector: Final result:', {
      isPropertyPaneOpen,
      foundSelector,
      foundElement: foundElement?.tagName,
      currentState: this.currentState
    });

    // Only trigger state change if the state actually changed
    if (isPropertyPaneOpen !== this.currentState) {
      console.log('🔍 PropertyPaneDetector: State changed from', this.currentState, 'to', isPropertyPaneOpen);
      console.log('🔍 Found via selector:', foundSelector);
      
      this.currentState = isPropertyPaneOpen;
      this.onStateChange(isPropertyPaneOpen);
    }
  }

  public addPropertyPaneStateClass(elements: NodeListOf<Element>, className: string): void {
    const isPropertyPaneOpen = this.isPropertyPaneOpen();
    
    elements.forEach((element: HTMLElement) => {
      if (isPropertyPaneOpen) {
        element.classList.add(className);
      } else {
        element.classList.remove(className);
      }
    });
  }

  public isPropertyPaneOpen(): boolean {
    const selectors = [
      '[data-automation-id="propertyPanePageCard"]',
      '.sp-PropertyPane:not([class*="Container"]):not(aside)',
      '[class*="propertyPane"]:not([class*="Container"]):not(aside)',
      '[class*="PropertyPane"]:not([class*="Container"]):not(aside)',
      '[data-automation-id*="propertyPane"]:not([data-automation-id*="Container"]):not(aside)',
      '[data-automation-id*="PropertyPane"]:not([data-automation-id*="Container"]):not(aside)',
      '.ms-Panel:not([class*="Container"]):not(aside)',
      '[class*="Panel"][class*="property"]:not([class*="Container"]):not(aside)',
      '[aria-label*="property" i][aria-label*="pane" i]:not([class*="Container"]):not(aside)'
    ];

    for (const selector of selectors) {
      const elements = document.querySelectorAll(selector);
      
      for (let i = 0; i < elements.length; i++) {
        const element = elements[i];
        const computedStyle = window.getComputedStyle(element);
        const rect = element.getBoundingClientRect();
        
        if (computedStyle.display !== 'none' && 
            computedStyle.visibility !== 'hidden' &&
            computedStyle.opacity !== '0') {
          if (rect.width > 0 && rect.height > 0) {
            // Additional check: property pane should be positioned on the right side of the screen
            const isOnRightSide = rect.left > (window.innerWidth * 0.6);
            const hasReasonableSize = rect.width > 250 && rect.height > 300;
            
            if (isOnRightSide && hasReasonableSize) {
              return true;
            }
          }
        }
      }
    }
    
    return false;
  }
}