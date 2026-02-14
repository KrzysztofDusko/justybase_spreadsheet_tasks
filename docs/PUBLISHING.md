# Publishing to npm

This document describes how to publish the `@justybase/spreadsheet-tasks` package to npm.

## Automated Publishing (Recommended)

The package is automatically published to npm when a new release is created on GitHub.

### Prerequisites

1. Ensure the `NPM_TOKEN` secret is set in the GitHub repository settings
2. The token must have publishing permissions for the `@justybase` scope

### Steps

1. Create a new tag: `git tag v1.0.0`
2. Push the tag: `git push origin v1.0.0`
3. Create a new release on GitHub using the tag
4. The publish workflow will automatically run and publish to npm

You can also manually trigger the publish workflow from the Actions tab in GitHub.

## Manual Publishing (Fallback)

If the automated workflow fails or you need to publish manually:

### Prerequisites

1. You must be logged in to npm: `npm login`
2. Your npm account must have access to publish under the `@justybase` scope
3. For first-time publishing, you may need to create the scope on npm first

### Steps

1. Build the project:
   ```bash
   npm run build
   ```

2. Run tests to ensure everything works:
   ```bash
   npm test
   ```

3. Publish to npm:
   ```bash
   npm publish --access public
   ```

## First Time Publishing

For the first time publishing a scoped package (`@justybase/spreadsheet-tasks`):

1. **Option A**: Create the `@justybase` organization on npm
   - Go to https://www.npmjs.com/
   - Create a new organization named `justybase`
   - Add your npm account as an owner

2. **Option B**: Publish as an unscoped package
   - Update `package.json` to change the name from `@justybase/spreadsheet-tasks` to `justybase-spreadsheet-tasks`
   - This doesn't require an organization

3. **Option C**: Use your personal npm scope
   - Update `package.json` to use your npm username as the scope
   - For example: `@yourusername/spreadsheet-tasks`

## Troubleshooting

### 404 Not Found Error

This typically means one of the following:

1. **First time publishing**: The scoped package doesn't exist yet and you need to create the scope/organization on npm
2. **Invalid token**: The `NPM_TOKEN` is expired or doesn't have the right permissions
3. **Scope access**: You don't have permission to publish under the `@justybase` scope

### Token Expired or Revoked

1. Create a new npm access token at https://www.npmjs.com/settings/[username]/tokens
2. Select "Automation" token type
3. Copy the token
4. Update the `NPM_TOKEN` secret in GitHub repository settings

### Permission Denied

Ensure your npm account:
1. Has 2FA enabled (required for publishing)
2. Is a member of the `@justybase` organization (if using scoped packages)
3. Has publishing permissions in the organization settings
