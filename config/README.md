# Configuration Files

This directory contains SPFx build and deployment configuration files.

## Local Development Configuration

### Using Your Tenant URL

To avoid having the placeholder `{tenantDomain}` in your development workflow, use the Microsoft-recommended environment variable approach:

1. **Copy the .env.example file:**
   ```bash
   cp .env.example .env
   ```

2. **Update with your tenant:**
   ```bash
   # Edit .env file
   SPFX_SERVE_TENANT_DOMAIN=yourtenant
   ```
   Replace `yourtenant` with your actual SharePoint tenant name (without `.sharepoint.com`).

3. **Run the dev server:**
   ```bash
   gulp serve
   ```
   SPFx will automatically use the `SPFX_SERVE_TENANT_DOMAIN` environment variable.

### Why This Approach?

- ✅ **Official Microsoft recommendation** - Uses standard SPFx environment variable
- ✅ **Git-safe** - `.env` is git-ignored (your tenant info stays private)
- ✅ **Simple** - Just set one environment variable
- ✅ **No file manipulation** - No need to copy or backup files
- ✅ **Works everywhere** - Supported by all SPFx versions

### How It Works

SPFx automatically replaces `{tenantDomain}` in `serve.json` with the value from `SPFX_SERVE_TENANT_DOMAIN`:

**serve.json:**
```json
{
  "initialPage": "https://{tenantDomain}.sharepoint.com/_layouts/workbench.aspx"
}
```

**With SPFX_SERVE_TENANT_DOMAIN=contoso, becomes:**
```
https://contoso.sharepoint.com/_layouts/workbench.aspx
```

## Configuration Files

| File | Purpose | Git Tracked |
|------|---------|-------------|
| `config.json` | Bundle configuration and entry points | ✅ Yes |
| `package-solution.json` | SharePoint package metadata | ✅ Yes |
| `serve.json` | Dev server config (with {tenantDomain} placeholder) | ✅ Yes |
| `deploy-azure-storage.json` | Azure CDN deployment (if used) | ❌ No (git-ignored) |
| `sass.json` | SASS compilation settings | ✅ Yes |
| `write-manifests.json` | Manifest generation settings | ✅ Yes |

## Tips

- **First-time setup:** Create `.env` file immediately after cloning
- **Team onboarding:** Include `.env.example` with your repo as a template
- **Multiple environments:** Use different `.env` files (`.env.dev`, `.env.prod`) and copy as needed
- **CI/CD:** Set `SPFX_SERVE_TENANT_DOMAIN` as a pipeline variable (not needed for builds, only for `gulp serve`)

## Reference

- [SPFx Documentation: Set up development environment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
