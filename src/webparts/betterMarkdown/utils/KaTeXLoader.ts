export class KaTeXLoader {
  private static isLoaded: boolean = false;

  public static loadKatexCss(): void {
    if (this.isLoaded) {
      return;
    }

    const head = document.getElementsByTagName('head')[0];
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = 'https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css';
    link.integrity = 'sha384-n8MVd4RsNIU0tAv4ct0nTaAbDJwPJzDEaqSD1odI+WdtXRGWt2kTvGFasHpSy3SV';
    link.crossOrigin = 'anonymous';
    head.appendChild(link);
    
    this.isLoaded = true;
  }

  public static get isKatexCssLoaded(): boolean {
    return this.isLoaded;
  }
}