# Better Markdown Styles Architecture

This directory contains modular SCSS files for the Better Markdown web part. The styles are split into logical components for easier maintenance and customization.

## File Structure

```
styles/
├── _variables.scss      # Color palette, spacing, typography, breakpoints
├── _mixins.scss         # Reusable mixins (responsive, cards, blockquotes)
├── _layout.scss         # Grid, layout, TOC sidebar positioning
├── _typography.scss     # Headings, paragraphs, links
├── _code.scss          # Code blocks, syntax highlighting
├── _blockquotes.scss   # Regular and styled blockquotes (info, warning, etc.)
├── _tables.scss        # Table styling with gradients and hover effects
├── _lists.scss         # Lists, grid-list, links-list
├── _toc.scss           # Table of contents styling
├── _mermaid.scss       # Mermaid diagram containers and math rendering
├── _editor.scss        # Editor mode layout and Monaco integration
└── _dark-theme.scss    # Dark theme overrides
```

## Customization Guide

### Colors

Edit `_variables.scss` to change the color scheme:

```scss
// Primary colors
$color-text-primary: #24292e;
$color-link: #0366d6;
$color-border: #e1e4e8;

// Blockquote variants
$color-info-bg: #e3f2fd;
$color-info-border: #2196f3;
```

### Spacing

All spacing uses variables for consistency:

```scss
$spacing-xs: 4px;   // Minimal spacing
$spacing-sm: 8px;   // Small spacing
$spacing-md: 16px;  // Medium spacing (default)
$spacing-lg: 24px;  // Large spacing
$spacing-xl: 32px;  // Extra large spacing
```

### Breakpoints

Responsive breakpoints are defined in `_variables.scss`:

```scss
$breakpoint-mobile: 768px;
$breakpoint-tablet: 992px;
$breakpoint-desktop: 1200px;
$breakpoint-wide: 1400px;
```

Use the responsive mixin in components:

```scss
.myComponent {
  width: 100%;

  @include respond-to(tablet) {
    width: 50%;
  }
}
```

### Adding New Components

1. Create a new partial: `_component-name.scss`
2. Add styles scoped under `.betterMarkdown`
3. Use variables from `_variables.scss`
4. Import in `BetterMarkdownWebPart.module.scss`:

```scss
@import './styles/component-name';
```

### Modifying Blockquotes

Blockquotes use a mixin for consistency. To add a new variant:

```scss
// In _blockquotes.scss
.blockquote:global(.is-custom) {
  @include styled-blockquote(
    $bg-color: #your-bg,
    $border-color: #your-border,
    $text-color: #your-text,
    $label: "🎨 Custom"
  );
}
```

### Dark Theme

Dark theme overrides are in `_dark-theme.scss`. Add component-specific dark styles:

```scss
&.dark {
  .myComponent {
    background-color: $color-dark-bg;
    color: $color-dark-text;
  }
}
```

## Benefits of This Architecture

1. **Maintainability** - Easy to find and update specific components
2. **Reusability** - Variables and mixins promote consistency
3. **Performance** - SCSS compiles to optimized CSS
4. **Scalability** - Easy to add new components without bloat
5. **Clarity** - Each file has a single responsibility

## Main Entry Point

The main file `BetterMarkdownWebPart.module.scss` imports all partials in the correct order:

```scss
// 1. Variables and utilities
@import './styles/variables';
@import './styles/mixins';

// 2. Components
@import './styles/layout';
@import './styles/typography';
// ... etc
```

## Compilation

The SCSS is compiled during the SPFx build process:
- Development: `gulp serve`
- Production: `gulp bundle --ship`

Compiled CSS includes CSS Modules hashing for scoping.
