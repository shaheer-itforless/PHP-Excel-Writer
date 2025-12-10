# Sharepoint Excel Writer (PHP Script)

## Prerequisites

* `php` 8.5

## Setup

```bash
composer install
```

### Example `.env` File

```txt
# Visit: Azure Portal > Microsoft Entra ID > Application Registrations > Search for your app
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here

# Visit: Azure Portal > Microsoft Entra ID > Application Registrations > Search for your app >
#        Manage > Certificates & secrets > Create secret
CLIENT_SECRET=your-client-secret-here

# Visit: https://{TenantName}.sharepoint.com/sites/{SiteName}/_api/site/id
# Copy : Edm.Guid
SITE_ID=your-site-id-here

# Visit: https://developer.microsoft.com/en-us/graph/graph-explorer
# GET:   https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives
# OR
# Get drive id and source file id using powerautomate
DRIVE_ID=your-drive-id-here
SOURCE_FILE_ID=your-file-path-here-in-encoded-strings

WORKSHEET_NAME=your-worksheet-name-here

# Random string generated for this script
API_SECRET=your-secret-token-here
```

### Example Config

Create a file in project root called `paths.json`.

`worksheetName` is optional. If not provided the value is `Sheet1`.

```json
{
    "libraryName": "/Shared Documents",
    "filePath": "/folder/file.xlsx",
    "worksheetName": "Sheet1"
}
```

## Run

```bash
php -S localhost:8000
```

In another terminal session

```bash
curl -X POST http://localhost:8000/excel_writer.php \
-H "Content-Type: application/json" \
-d '{
"secret": "your-secret-token",
"FirstName": "FirstName",
"LastName": "LastName",
"Email": "example@email.com"
}'
```
