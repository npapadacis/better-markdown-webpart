'use strict';

const build = require('@microsoft/sp-build-web');

// Load environment variables from .env file if it exists
// This allows setting SPFX_SERVE_TENANT_DOMAIN without committing it to Git
require('dotenv').config();

build.addSuppression(/Warning/gi);

// Display tenant info if configured
if (process.env.SPFX_SERVE_TENANT_DOMAIN) {
  console.log(`✓ Using tenant: ${process.env.SPFX_SERVE_TENANT_DOMAIN}`);
  console.log(`  Opening: https://${process.env.SPFX_SERVE_TENANT_DOMAIN}.sharepoint.com/_layouts/workbench.aspx`);
} else {
  console.log('ℹ No tenant configured. Using placeholder {tenantDomain} in workbench URL');
  console.log('  Tip: Create a .env file with SPFX_SERVE_TENANT_DOMAIN=yourtenant');
}

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
