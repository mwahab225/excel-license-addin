# Excel License Add-in (Revised)

## What It Does
- License validation UI (Company Name, User Name, Key)
- Calls your API to check the license key
- Stores result in Excel settings
- Collects worksheet names and outputs them in a "Summary" sheet

## Setup Instructions
1. Extract this ZIP to: `C:/OfficeAddins/LicenseValidator/`
2. Share the folder over the network as: `\PC-049\OfficeAddins\LicenseValidator\`
3. Add this path to Excel's Trust Center (Trusted Add-in Catalogs)
4. Restart Excel → Go to **Insert > My Add-ins > Shared Folder**

## API Endpoint Example

POST `https://your-api-domain.com/validate-license`
```json
{
  "company": "AE Core",
  "user": "Ali",
  "licenseKey": "ABCD-1234"
}
```

Returns:
```json
{ "status": "valid" }
```

Edit `taskpane.js` with your API domain.
