# Configuration Files

This directory contains SPFx build and deployment configuration files.

## Local Development Configuration

### Using Your Tenant URL

To avoid having the placeholder `{tenant-name}` in your development workflow:

1. **Copy the example file:**
   ```bash
   cp serve.json serve.local.json
   ```

2. **Update with your tenant:**
   ```json
   {
     "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
     "port": 4321,
     "https": true,
     "initialPage": "https://{tenant-name}.sharepoint.com/_layouts/workbench.aspx"
   }
   ```
   Replace `{tenant-name}` with your actual SharePoint tenant name.

3. **Run the dev server:**
   ```bash
   gulp serve
   ```
   SPFx will automatically use `serve.local.json` instead of `serve.json`.

### Why This Approach?

- ✅ `serve.local.json` is git-ignored (your tenant info stays private)
- ✅ `serve.json` stays with the placeholder (safe to commit)
- ✅ SPFx automatically prioritizes `serve.local.json` when it exists

## Configuration Files

| File | Purpose | Git Tracked |
|------|---------|-------------|
| `config.json` | Bundle configuration and entry points | ✅ Yes |
| `package-solution.json` | SharePoint package metadata | ✅ Yes |
| `serve.json` | Dev server config (with placeholder) | ✅ Yes |
| `serve.local.json` | Dev server config (with real tenant) | ❌ No (git-ignored) |
| `deploy-azure-storage.json` | Azure CDN deployment (if used) | ❌ No (git-ignored) |
| `sass.json` | SASS compilation settings | ✅ Yes |
| `write-manifests.json` | Manifest generation settings | ✅ Yes |

## Tips

- **First-time setup:** Copy `serve.json` to `serve.local.json` immediately after cloning
- **Team onboarding:** Include instructions in your team wiki/docs to create `serve.local.json`
- **Multiple tenants:** Create multiple files like `serve.local.dev.json`, `serve.local.prod.json` and copy as needed
- **CI/CD:** Use `serve.json` (with placeholder) in automated builds since they don't need the dev server
